# OneDrive Duplicate Photo Finder

Find and safely recycle exact duplicate photos in your personal OneDrive using Microsoft Graph. No downloads, dry-run by default, and it never deletes anything permanently.

Two builds, one behavior:

- `OneDrive-DuplicatePhotoFinder.ps1` - PowerShell (Windows, macOS, Linux via PowerShell 7).
- `onedrive_duplicate_finder.py` - Python 3.8+ (Windows, macOS, Linux, and a Chromebook's Linux container). Standard library only, no pip install.

## Why this exists

Neither OneDrive nor Google Photos ships a duplicate photo finder, and people have asked for one for years. OneDrive in particular creates `(1)` copies and re-uploads the same files across devices. This tool finds the byte-identical copies and moves the extras to the Recycle Bin, where they are recoverable.

## How it works

Microsoft Graph already stores a content hash (`quickXorHash`) for every file in OneDrive. This tool reads those hashes, groups files by hash plus size, keeps one copy per group per your chosen strategy, and reports or recycles the rest. Because it compares hashes Graph already has, it never downloads your photos.

## Safety

- **Dry-run by default.** It reports what it would do and changes nothing unless you pass `-Execute` (PowerShell) or `--execute` (Python).
- **Recycle Bin only.** A normal Graph delete is recoverable for about 30 days. This tool never calls permanent delete.
- **Confirmation gate.** Before moving anything it asks you to type `YES` (skippable with `-Force` / `--force` for unattended runs).
- **CSV report** of every duplicate set and the planned action, written before any change.

## Known limitation

This finds **exact** duplicates only: byte-identical files. It does not yet find visually similar shots or the same photo resized or re-encoded, which requires downloading and perceptual hashing. That is a planned next phase.

## Requirements

- **PowerShell build:** PowerShell 5.1+ (or 7+ for macOS/Linux) and the `Microsoft.Graph.Authentication` module:
  ```
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```
- **Python build:** Python 3.8+ and an Entra (Azure AD) app registration you create once:
  - Type: Public client / native (mobile & desktop).
  - Allow public client / device code flow: Yes.
  - Delegated permission: Microsoft Graph -> `Files.ReadWrite`.

  Then set the client ID:
  ```
  export DARKPULSE_MS_CLIENT_ID="your-client-id"     # macOS / Linux / Chromebook
  $env:DARKPULSE_MS_CLIENT_ID = "your-client-id"     # Windows PowerShell
  ```
  The PowerShell build does not need this; `Connect-MgGraph` uses Microsoft's built-in client.

## Usage

PowerShell:
```powershell
.\OneDrive-DuplicatePhotoFinder.ps1                       # dry run
.\OneDrive-DuplicatePhotoFinder.ps1 -Execute              # recycle extras (asks to confirm)
.\OneDrive-DuplicatePhotoFinder.ps1 -KeepStrategy KeepNewest -IncludeVideos
```

Python:
```bash
python onedrive_duplicate_finder.py                       # dry run
python onedrive_duplicate_finder.py --execute             # recycle extras (asks to confirm)
python onedrive_duplicate_finder.py --keep newest --include-videos
```

Which copy is kept: oldest (default), newest, or the one with the shortest path.

## Chromebook note

Plain ChromeOS cannot run local scripts. Enable the Linux (Crostini) container in Settings, open the Linux terminal, install Python 3, then run the Python build as above.

## Roadmap

- Near-duplicate detection (similar and re-encoded photos) via perceptual hashing.
- A hosted web version so no local setup or app registration is needed.
- Google Photos support is blocked for now: Google restricted its Photos API so third-party apps can no longer scan a full library.

## License

MIT (add a `LICENSE` file). Use at your own risk; review the dry-run report before running with execute.
