#Requires -Version 5.1
<#
.SYNOPSIS
    Finds EXACT duplicate photos (and optionally videos) in your personal
    OneDrive using Microsoft Graph, and can move the extra copies to the
    OneDrive Recycle Bin. Dry-run by default. Never permanently deletes.

.DESCRIPTION
    Reads file content hashes (quickXorHash) that Microsoft Graph already
    stores for every OneDrive item, so it detects duplicates WITHOUT
    downloading any files. Groups items by hash + size, keeps one copy per
    group per your chosen strategy, and reports (or, with -Execute, recycles)
    the rest.

.NOTES
    SCRIPT HIGHLIGHTS
      - Exact-duplicate detection via Graph quickXorHash. No downloads.
      - Dry-run by default. -Execute is required to move anything.
      - Recycle Bin only. A normal Graph DELETE is recoverable (~30 days).
        This script never calls permanentDelete.
      - Timestamped CSV report of every duplicate set and the planned action.
      - Confirmation gate before any move, unless -Force is passed.

    REQUIREMENTS
      - Windows PowerShell 5.1+ or PowerShell 7+.
      - Module: Microsoft.Graph.Authentication
          Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
      - Signs in as YOU (delegated auth, Files.ReadWrite). Works on your own
        OneDrive. Serving other users would need an Entra app registration.

    KNOWN LIMITATION (read first)
      - This finds EXACT duplicates only: byte-identical files, which is the
        common "(1)" copy and re-upload case. It does NOT find visually
        similar shots or the same photo resized/re-encoded; that needs
        downloading plus perceptual hashing (a later phase).

    DISCLAIMER
      Provided as-is. Review the dry-run report before running -Execute.
      Moved items sit in the OneDrive Recycle Bin and can be restored.

    Last Updated: 2026-08-23

.PARAMETER KeepStrategy
    Which copy in each duplicate set to KEEP. The others are recycled.
    KeepOldest (default), KeepNewest, or KeepShortestPath.

.PARAMETER IncludeVideos
    Also scan video files, not just images.

.PARAMETER Execute
    Actually move the extra copies to the Recycle Bin. Omit for a dry run.

.PARAMETER Force
    Skip the interactive confirmation prompt when using -Execute (for
    unattended runs). Use with care.

.PARAMETER ExportPath
    CSV report path. Defaults to a timestamped file in the current folder.

.PARAMETER OpenReport
    Open the CSV report when finished.

.PARAMETER KeepConnected
    Do not disconnect the Graph session at the end (useful for batch runs).

.EXAMPLE
    .\OneDrive-DuplicatePhotoFinder.ps1
    Dry run: report duplicates, change nothing.

.EXAMPLE
    .\OneDrive-DuplicatePhotoFinder.ps1 -Execute -KeepStrategy KeepNewest
    Recycle all but the newest copy in each duplicate set (after confirming).
#>

[CmdletBinding()]
param(
    [ValidateSet('KeepOldest', 'KeepNewest', 'KeepShortestPath')]
    [string] $KeepStrategy = 'KeepOldest',

    [switch] $IncludeVideos,

    [switch] $Execute,

    [switch] $Force,

    [string] $ExportPath = ("OneDrive-Duplicates_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date)),

    [switch] $OpenReport,

    [switch] $KeepConnected
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Extensions treated as images / videos when Graph does not supply a mimeType.
$script:ImageExt = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff',
                     '.heic', '.heif', '.webp', '.raw', '.dng', '.cr2', '.nef')
$script:VideoExt = @('.mp4', '.mov', '.m4v', '.avi', '.mkv', '.wmv', '.3gp', '.hevc')


function Write-Status {
    <# Colored, consistent status line. #>
    param(
        [Parameter(Mandatory)][string] $Message,
        [ValidateSet('Info', 'Good', 'Warn', 'Bad')][string] $Level = 'Info'
    )
    $color = switch ($Level) {
        'Good' { 'Green' }
        'Warn' { 'Yellow' }
        'Bad'  { 'Red' }
        default { 'Cyan' }
    }
    Write-Host $Message -ForegroundColor $color
}


function Test-GraphModule {
    <# Ensure the Graph auth module is available, with install guidance. #>
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        Write-Status "Required module 'Microsoft.Graph.Authentication' is not installed." 'Bad'
        Write-Status "Install it, then re-run:" 'Warn'
        Write-Status "  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" 'Warn'
        return $false
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    return $true
}


function Connect-Graph {
    <# Delegated sign-in with the least scope needed to read and recycle. #>
    try {
        Connect-MgGraph -Scopes 'Files.ReadWrite' -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        if (-not $ctx) { throw "No Graph context after connect." }
        Write-Status "Signed in as: $($ctx.Account)" 'Good'
        return $true
    }
    catch {
        Write-Status "Graph sign-in failed: $($_.Exception.Message)" 'Bad'
        return $false
    }
}


function Get-Value {
    <# Null-safe nested lookup for the hashtables Graph returns. #>
    param([object] $Object, [string[]] $Path)
    $current = $Object
    foreach ($key in $Path) {
        if ($null -eq $current) { return $null }
        if ($current -is [System.Collections.IDictionary] -and $current.Contains($key)) {
            $current = $current[$key]
        }
        else {
            return $null
        }
    }
    return $current
}


function Test-IsMediaFile {
    <# Decide whether a Graph file item is a photo (or video, if requested). #>
    param([object] $Item)

    $mime = Get-Value $Item @('file', 'mimeType')
    $name = [string](Get-Value $Item @('name'))
    $ext = ([System.IO.Path]::GetExtension($name)).ToLowerInvariant()

    $isImage = ($mime -like 'image/*') -or ($script:ImageExt -contains $ext)
    if ($isImage) { return $true }

    if ($IncludeVideos) {
        $isVideo = ($mime -like 'video/*') -or ($script:VideoExt -contains $ext)
        if ($isVideo) { return $true }
    }
    return $false
}


function Get-OneDriveMedia {
    <#
        Recursively enumerate the signed-in user's OneDrive and return one
        PSCustomObject per media file, including its content hash.
    #>
    $results = New-Object System.Collections.Generic.List[object]
    $noHash = 0
    $seen = 0

    # Breadth-first walk of folders starting at the drive root.
    $queue = New-Object System.Collections.Generic.Queue[string]
    $queue.Enqueue('/me/drive/root/children')

    while ($queue.Count -gt 0) {
        $uri = $queue.Dequeue()

        while ($uri) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            $items = @()
            if ($resp.Contains('value')) { $items = $resp['value'] }

            foreach ($item in $items) {
                if (Get-Value $item @('folder')) {
                    # Descend into subfolders.
                    $id = Get-Value $item @('id')
                    if ($id) { $queue.Enqueue("/me/drive/items/$id/children") }
                    continue
                }

                if (-not (Get-Value $item @('file'))) { continue }   # not a file
                if (-not (Test-IsMediaFile $item)) { continue }      # not media

                $seen++
                if (($seen % 200) -eq 0) {
                    Write-Progress -Activity 'Scanning OneDrive' -Status "$seen media files found"
                }

                $hash = Get-Value $item @('file', 'hashes', 'quickXorHash')
                if (-not $hash) {
                    $noHash++
                    continue   # cannot compare without a hash; report count later
                }

                $parentPath = [string](Get-Value $item @('parentReference', 'path'))
                $results.Add([pscustomobject]@{
                    Id       = [string](Get-Value $item @('id'))
                    Name     = [string](Get-Value $item @('name'))
                    Path     = "$parentPath/$([string](Get-Value $item @('name')))"
                    Size     = [long](Get-Value $item @('size'))
                    Hash     = [string]$hash
                    Created  = [string](Get-Value $item @('fileSystemInfo', 'createdDateTime'))
                    Modified = [string](Get-Value $item @('fileSystemInfo', 'lastModifiedDateTime'))
                })
            }

            # Follow pagination within this folder listing.
            $uri = if ($resp.Contains('@odata.nextLink')) { $resp['@odata.nextLink'] } else { $null }
        }
    }

    Write-Progress -Activity 'Scanning OneDrive' -Completed
    return [pscustomobject]@{
        Media  = $results
        NoHash = $noHash
        Total  = $seen
    }
}


function Select-Keeper {
    <# Choose the item to KEEP from a duplicate set, per the strategy. #>
    param([object[]] $Set)
    switch ($KeepStrategy) {
        'KeepNewest'      { return ($Set | Sort-Object Modified -Descending | Select-Object -First 1) }
        'KeepShortestPath' { return ($Set | Sort-Object { $_.Path.Length } | Select-Object -First 1) }
        default           { return ($Set | Sort-Object Created | Select-Object -First 1) }  # KeepOldest
    }
}


# ---- Main ------------------------------------------------------------------

$exitCode = 0
$connected = $false
try {
    Write-Status "OneDrive Duplicate Photo Finder  (mode: $(if ($Execute) {'EXECUTE'} else {'DRY RUN'}))" 'Info'

    if (-not (Test-GraphModule)) { exit 1 }
    if (-not (Connect-Graph)) { exit 1 }
    $connected = $true

    Write-Status "Scanning OneDrive for media files (this can take a while on large libraries)..." 'Info'
    $scan = Get-OneDriveMedia
    Write-Status "Media files with a usable hash: $($scan.Media.Count) (scanned $($scan.Total); $($scan.NoHash) had no hash and were skipped)." 'Info'

    if ($scan.Media.Count -eq 0) {
        Write-Status "No comparable media found. Nothing to do." 'Warn'
        exit 0
    }

    # Group by content hash AND size (size guards against any hash collision).
    # Wrap in @() so .Count means "number of duplicate sets" even when there is
    # exactly one set (a bare GroupInfo would report its member count instead).
    $groups = @($scan.Media | Group-Object -Property { "$($_.Hash)|$($_.Size)" } |
        Where-Object { $_.Count -gt 1 })

    if ($groups.Count -eq 0) {
        Write-Status "No exact duplicates found. Your library is clean on that measure." 'Good'
        exit 0
    }

    $reportRows = New-Object System.Collections.Generic.List[object]
    [long] $reclaimBytes = 0
    [int]  $dupCount = 0

    foreach ($g in $groups) {
        $set = @($g.Group)
        $keeper = Select-Keeper -Set $set
        foreach ($item in $set) {
            $isKeeper = ($item.Id -eq $keeper.Id)
            if (-not $isKeeper) {
                $dupCount++
                $reclaimBytes += $item.Size
            }
            $reportRows.Add([pscustomobject]@{
                Action   = if ($isKeeper) { 'KEEP' } else { if ($Execute) { 'RECYCLED' } else { 'WOULD-RECYCLE' } }
                Name     = $item.Name
                Path     = $item.Path
                SizeMB   = [math]::Round($item.Size / 1MB, 2)
                Created  = $item.Created
                Hash     = $item.Hash
                Id       = $item.Id
            })
        }
    }

    $reclaimMB = [math]::Round($reclaimBytes / 1MB, 2)
    Write-Status "Duplicate sets: $($groups.Count).  Extra copies: $dupCount.  Reclaimable: $reclaimMB MB." 'Info'

    # Write the CSV report before taking any action.
    $reportRows | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Status "Report written: $ExportPath" 'Good'

    # Show a small sample so the user can sanity-check before executing.
    Write-Host ""
    $reportRows | Where-Object { $_.Action -ne 'KEEP' } |
        Select-Object -First 10 Action, Name, SizeMB, Path |
        Format-Table -AutoSize | Out-Host

    if (-not $Execute) {
        Write-Status "DRY RUN complete. Would recycle $dupCount file(s), reclaiming ~$reclaimMB MB." 'Warn'
        Write-Status "Re-run with -Execute to move the extra copies to the Recycle Bin." 'Warn'
        if ($OpenReport) { Invoke-Item -Path $ExportPath }
        exit 0
    }

    # Destructive path: confirm unless -Force.
    if (-not $Force) {
        Write-Host ""
        Write-Status "About to move $dupCount file(s) to the OneDrive Recycle Bin (recoverable)." 'Warn'
        $answer = Read-Host "Type YES to proceed"
        if ($answer -ne 'YES') {
            Write-Status "Aborted by user. Nothing was moved." 'Warn'
            exit 0
        }
    }

    $moved = 0
    $failed = 0
    $toMove = @($reportRows | Where-Object { $_.Action -eq 'RECYCLED' })
    $i = 0
    foreach ($row in $toMove) {
        $i++
        Write-Progress -Activity 'Moving duplicates to Recycle Bin' `
            -Status "$i of $($toMove.Count): $($row.Name)" `
            -PercentComplete (($i / [math]::Max($toMove.Count, 1)) * 100)
        try {
            # A normal DELETE moves the item to the Recycle Bin (recoverable).
            Invoke-MgGraphRequest -Method DELETE -Uri "/me/drive/items/$($row.Id)" -ErrorAction Stop | Out-Null
            $moved++
            Write-Status "  Recycled: $($row.Name)" 'Good'
        }
        catch {
            $failed++
            Write-Status "  FAILED:   $($row.Name) -> $($_.Exception.Message)" 'Bad'
        }
    }
    Write-Progress -Activity 'Moving duplicates to Recycle Bin' -Completed

    Write-Status "Done. Moved $moved to Recycle Bin. Failed: $failed. Report: $ExportPath" $(if ($failed) { 'Warn' } else { 'Good' })
    if ($OpenReport) { Invoke-Item -Path $ExportPath }
    if ($failed -gt 0) { $exitCode = 1 }
}
catch {
    Write-Status "Unhandled error: $($_.Exception.Message)" 'Bad'
    $exitCode = 1
}
finally {
    if ($connected -and -not $KeepConnected) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
}

exit $exitCode
