#!/usr/bin/env python3
"""
OneDrive Duplicate Photo Finder - cross-platform (Windows / macOS / Linux /
ChromeOS Linux container)
==============================================================================
SCRIPT HIGHLIGHTS
  - Finds EXACT duplicate photos (and optionally videos) in your personal
    OneDrive via Microsoft Graph, using the content hash Graph already stores
    (quickXorHash). No files are downloaded.
  - Dry-run by default. --execute is required to move anything.
  - Recycle Bin only. A normal Graph DELETE is recoverable (~30 days). This
    script never permanently deletes.
  - One codebase runs anywhere Python 3.8+ runs. Standard library only, no pip.

REQUIREMENTS
  - Python 3.8+.
  - An Entra (Azure AD) app registration you create once:
      * Type: Public client / native (mobile & desktop).
      * Allow public client / device code flow: Yes.
      * Delegated permission: Microsoft Graph -> Files.ReadWrite.
    Then set the app's Application (client) ID:
      Windows PowerShell:  $env:DARKPULSE_MS_CLIENT_ID = "your-client-id"
      macOS/Linux/Chrome:  export DARKPULSE_MS_CLIENT_ID="your-client-id"
    (Or pass --client-id on the command line.)

  Why an app registration: device-code sign-in needs a client ID. This is a
  one-time setup and is the same registration you would need to offer the tool
  to other users.

KNOWN LIMITATION (read first)
  - Finds EXACT duplicates only (byte-identical files: the common "(1)" copy
    and re-upload case). It does NOT find visually similar shots or the same
    photo resized/re-encoded; that needs downloading plus perceptual hashing.
  - Authenticates as YOU, against your own OneDrive.

USAGE
  python onedrive_duplicate_finder.py                 # dry run, change nothing
  python onedrive_duplicate_finder.py --execute       # move extras to Recycle Bin
  python onedrive_duplicate_finder.py --keep newest --include-videos

PLATFORM SETUP
  - Windows/macOS/Linux: install Python 3, then run as above.
  - Chromebook: enable the Linux (Crostini) container in Settings, open the
    Linux terminal, install Python 3, then run as above. Plain ChromeOS with no
    Linux container cannot run local scripts.

DISCLAIMER
  Provided as-is. Review the dry-run report before running --execute. Moved
  items go to the OneDrive Recycle Bin and can be restored.

Last Updated: 2026-08-23
==============================================================================
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, List, Optional, Tuple

# ---- Configuration ---------------------------------------------------------

AUTHORITY = "https://login.microsoftonline.com/common"
DEVICE_CODE_URL = f"{AUTHORITY}/oauth2/v2.0/devicecode"
TOKEN_URL = f"{AUTHORITY}/oauth2/v2.0/token"
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = "Files.ReadWrite offline_access openid profile"

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff",
             ".heic", ".heif", ".webp", ".raw", ".dng", ".cr2", ".nef"}
VIDEO_EXT = {".mp4", ".mov", ".m4v", ".avi", ".mkv", ".wmv", ".3gp", ".hevc"}


# ---- Small HTTP helpers (stdlib only) --------------------------------------

def _http(url: str, method: str = "GET", data: Optional[bytes] = None,
          headers: Optional[Dict[str, str]] = None, timeout: int = 60
          ) -> Tuple[int, Dict[str, Any]]:
    """Perform an HTTP request; return (status, parsed_json_or_empty)."""
    req = urllib.request.Request(url, data=data, method=method,
                                 headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8") if resp.length != 0 else ""
            return resp.status, (json.loads(body) if body else {})
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8")
        except Exception:
            pass
        parsed = {}
        try:
            parsed = json.loads(detail) if detail else {}
        except Exception:
            parsed = {"raw": detail}
        return exc.code, parsed
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Network error: {exc.reason}") from exc


# ---- Authentication (device code flow) -------------------------------------

def device_code_login(client_id: str) -> str:
    """Run the device-code flow and return an access token."""
    body = urllib.parse.urlencode({"client_id": client_id,
                                   "scope": SCOPES}).encode("utf-8")
    status, resp = _http(DEVICE_CODE_URL, "POST", body,
                         {"Content-Type": "application/x-www-form-urlencoded"})
    if status != 200 or "device_code" not in resp:
        raise RuntimeError(f"Device code request failed: {resp}")

    print("\n" + "=" * 60)
    print("To sign in, open this URL in a browser:")
    print("  " + resp.get("verification_uri", "https://microsoft.com/devicelogin"))
    print("and enter this code:")
    print("  " + resp.get("user_code", ""))
    print("=" * 60 + "\n")

    device_code = resp["device_code"]
    interval = int(resp.get("interval", 5))
    # Poll until the user completes sign-in (or it expires).
    while True:
        time.sleep(interval)
        poll_body = urllib.parse.urlencode({
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id,
            "device_code": device_code,
        }).encode("utf-8")
        status, tok = _http(TOKEN_URL, "POST", poll_body,
                            {"Content-Type": "application/x-www-form-urlencoded"})
        if status == 200 and "access_token" in tok:
            return tok["access_token"]
        err = tok.get("error", "")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        raise RuntimeError(f"Sign-in failed: {tok.get('error_description', tok)}")


# ---- Graph calls -----------------------------------------------------------

def graph_get(path_or_url: str, token: str) -> Dict[str, Any]:
    url = path_or_url if path_or_url.startswith("http") else GRAPH + path_or_url
    status, resp = _http(url, "GET", None, {"Authorization": f"Bearer {token}"})
    if status != 200:
        raise RuntimeError(f"Graph GET {url} failed ({status}): {resp}")
    return resp


def graph_delete(item_id: str, token: str) -> None:
    """DELETE moves the item to the Recycle Bin (recoverable)."""
    url = f"{GRAPH}/me/drive/items/{item_id}"
    status, resp = _http(url, "DELETE", None, {"Authorization": f"Bearer {token}"})
    if status not in (200, 204):
        raise RuntimeError(f"Delete failed ({status}): {resp}")


# ---- Media enumeration -----------------------------------------------------

def _ext(name: str) -> str:
    dot = name.rfind(".")
    return name[dot:].lower() if dot >= 0 else ""


def is_media(item: Dict[str, Any], include_videos: bool) -> bool:
    file_facet = item.get("file") or {}
    mime = (file_facet.get("mimeType") or "").lower()
    ext = _ext(item.get("name", ""))
    if mime.startswith("image/") or ext in IMAGE_EXT:
        return True
    if include_videos and (mime.startswith("video/") or ext in VIDEO_EXT):
        return True
    return False


def enumerate_media(token: str, include_videos: bool) -> Tuple[List[Dict[str, Any]], int, int]:
    """Walk OneDrive; return (media_records, scanned_count, no_hash_count)."""
    media: List[Dict[str, Any]] = []
    scanned = 0
    no_hash = 0
    queue: List[str] = ["/me/drive/root/children"]

    while queue:
        url = queue.pop(0)
        while url:
            resp = graph_get(url, token)
            for item in resp.get("value", []):
                if item.get("folder"):
                    item_id = item.get("id")
                    if item_id:
                        queue.append(f"/me/drive/items/{item_id}/children")
                    continue
                if not item.get("file"):
                    continue
                if not is_media(item, include_videos):
                    continue
                scanned += 1
                if scanned % 200 == 0:
                    print(f"  ...scanned {scanned} media files", flush=True)
                hashes = (item.get("file") or {}).get("hashes") or {}
                qhash = hashes.get("quickXorHash")
                if not qhash:
                    no_hash += 1
                    continue
                parent = (item.get("parentReference") or {}).get("path", "")
                fsi = item.get("fileSystemInfo") or {}
                media.append({
                    "id": item.get("id", ""),
                    "name": item.get("name", ""),
                    "path": f"{parent}/{item.get('name', '')}",
                    "size": int(item.get("size", 0)),
                    "hash": qhash,
                    "created": fsi.get("createdDateTime", ""),
                    "modified": fsi.get("lastModifiedDateTime", ""),
                })
            url = resp.get("@odata.nextLink")
    return media, scanned, no_hash


# ---- Duplicate logic -------------------------------------------------------

def pick_keeper(group: List[Dict[str, Any]], strategy: str) -> Dict[str, Any]:
    """Choose which copy to keep; the rest are recycled."""
    if strategy == "newest":
        return sorted(group, key=lambda r: r["modified"], reverse=True)[0]
    if strategy == "shortest":
        return sorted(group, key=lambda r: len(r["path"]))[0]
    return sorted(group, key=lambda r: r["created"])[0]  # oldest (default)


def group_duplicates(media: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
    """Group by hash + size; return only groups with more than one member."""
    buckets: Dict[str, List[Dict[str, Any]]] = {}
    for rec in media:
        key = f"{rec['hash']}|{rec['size']}"
        buckets.setdefault(key, []).append(rec)
    return [g for g in buckets.values() if len(g) > 1]


# ---- Main ------------------------------------------------------------------

def run(args: argparse.Namespace) -> int:
    client_id = args.client_id or os.environ.get("DARKPULSE_MS_CLIENT_ID", "").strip()
    if not client_id:
        sys.stderr.write(
            "ERROR: no client ID. Set DARKPULSE_MS_CLIENT_ID or pass --client-id.\n"
            "See the REQUIREMENTS section at the top of this file.\n")
        return 1

    mode = "EXECUTE" if args.execute else "DRY RUN"
    print(f"OneDrive Duplicate Photo Finder  (mode: {mode})")

    try:
        token = device_code_login(client_id)
    except RuntimeError as exc:
        sys.stderr.write(f"Sign-in error: {exc}\n")
        return 1

    print("Signed in. Scanning OneDrive (this can take a while on large libraries)...")
    try:
        media, scanned, no_hash = enumerate_media(token, args.include_videos)
    except RuntimeError as exc:
        sys.stderr.write(f"Scan error: {exc}\n")
        return 1

    print(f"Media with a usable hash: {len(media)} "
          f"(scanned {scanned}; {no_hash} had no hash and were skipped).")
    if not media:
        print("Nothing comparable found. Done.")
        return 0

    groups = group_duplicates(media)
    if not groups:
        print("No exact duplicates found. Your library is clean on that measure.")
        return 0

    rows: List[Dict[str, Any]] = []
    reclaim = 0
    dup_count = 0
    for g in groups:
        keeper = pick_keeper(g, args.keep)
        for rec in g:
            is_keeper = rec["id"] == keeper["id"]
            if not is_keeper:
                dup_count += 1
                reclaim += rec["size"]
            rows.append({
                "action": "KEEP" if is_keeper else ("RECYCLED" if args.execute else "WOULD-RECYCLE"),
                "name": rec["name"],
                "path": rec["path"],
                "size_mb": round(rec["size"] / (1024 * 1024), 2),
                "created": rec["created"],
                "hash": rec["hash"],
                "id": rec["id"],
            })

    reclaim_mb = round(reclaim / (1024 * 1024), 2)
    print(f"Duplicate sets: {len(groups)}.  Extra copies: {dup_count}.  "
          f"Reclaimable: {reclaim_mb} MB.")

    # Write the CSV report before any action.
    try:
        with open(args.export, "w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
            writer.writeheader()
            writer.writerows(rows)
        print(f"Report written: {args.export}")
    except OSError as exc:
        sys.stderr.write(f"Could not write report: {exc}\n")

    # Show a small sample.
    print("\nSample (first 10 extras that would be recycled):")
    shown = 0
    for r in rows:
        if r["action"] != "KEEP":
            print(f"  {r['action']}  {r['size_mb']:>7} MB  {r['path']}")
            shown += 1
            if shown >= 10:
                break

    if not args.execute:
        print(f"\nDRY RUN complete. Would recycle {dup_count} file(s), "
              f"reclaiming ~{reclaim_mb} MB.")
        print("Re-run with --execute to move the extra copies to the Recycle Bin.")
        return 0

    # Destructive path: confirm unless --force.
    if not args.force:
        print(f"\nAbout to move {dup_count} file(s) to the OneDrive Recycle Bin "
              f"(recoverable).")
        if input("Type YES to proceed: ").strip() != "YES":
            print("Aborted. Nothing was moved.")
            return 0

    moved = 0
    failed = 0
    to_move = [r for r in rows if r["action"] == "RECYCLED"]
    for i, r in enumerate(to_move, 1):
        try:
            graph_delete(r["id"], token)
            moved += 1
            print(f"  [{i}/{len(to_move)}] Recycled: {r['name']}")
        except RuntimeError as exc:
            failed += 1
            print(f"  [{i}/{len(to_move)}] FAILED: {r['name']} -> {exc}")

    print(f"\nDone. Moved {moved} to Recycle Bin. Failed: {failed}. Report: {args.export}")
    return 1 if failed else 0


def main() -> None:
    ts = time.strftime("%Y%m%d_%H%M%S")
    parser = argparse.ArgumentParser(
        description="Find exact duplicate photos in OneDrive (dry-run by default).")
    parser.add_argument("--execute", action="store_true",
                        help="Move extra copies to the Recycle Bin (default: dry run).")
    parser.add_argument("--keep", choices=["oldest", "newest", "shortest"],
                        default="oldest", help="Which copy to keep in each set.")
    parser.add_argument("--include-videos", action="store_true",
                        help="Also scan video files.")
    parser.add_argument("--force", action="store_true",
                        help="Skip the confirmation prompt with --execute.")
    parser.add_argument("--client-id", default="",
                        help="Entra app (client) ID. Overrides DARKPULSE_MS_CLIENT_ID.")
    parser.add_argument("--export", default=f"OneDrive-Duplicates_{ts}.csv",
                        help="CSV report path.")
    args = parser.parse_args()
    sys.exit(run(args))


if __name__ == "__main__":
    main()
