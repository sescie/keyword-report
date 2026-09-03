"""
drive_utils.py — read a month's screenshots from a Google Drive folder,
and write the finished report back into that same folder. Uses the
Drive API v3 directly (via the credentials from drive_auth.py) — no
local mounting, no service account, just the signed-in person's own
Drive access for exactly the folders they choose to point at.
"""

import io
import os

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

FOLDER_MIME = "application/vnd.google-apps.folder"


def get_drive_service(credentials):
    return build("drive", "v3", credentials=credentials)


def extract_folder_id(url_or_id: str) -> str:
    """Accepts a full Drive folder URL (several formats Google actually
    uses) or a bare folder ID, and returns just the ID either way — so
    a person can literally paste whatever's in their browser's address
    bar without needing to know what a "folder ID" even is."""
    s = url_or_id.strip()
    if "/folders/" in s:
        tail = s.split("/folders/", 1)[1]
        return tail.split("?", 1)[0].split("/", 1)[0]
    if "id=" in s:
        tail = s.split("id=", 1)[1]
        return tail.split("&", 1)[0]
    return s  # already looks like a bare ID


def list_children(service, folder_id: str) -> list[dict]:
    """All direct children (files and subfolders) of a Drive folder,
    handling pagination — a month's worth of screenshots can easily
    exceed the API's single-page result limit."""
    items = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token, pageSize=1000,
        ).execute()
        items.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return items


def download_file(service, file_id: str, local_path: str) -> None:
    """Downloads a single Drive file (e.g. a PPTX template) directly to
    a local path — same underlying mechanism as download_folder_tree,
    just for one already-known file ID rather than a whole folder."""
    request = service.files().get_media(fileId=file_id)
    with io.FileIO(local_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def download_folder_tree(service, folder_id: str, local_dir: str, _depth=0) -> int:
    """Recursively mirrors a Drive folder into a local directory,
    preserving the exact folder structure (so the existing, unchanged
    scan_and_detect logic can run against the downloaded copy exactly
    as it would against a local drive). Returns the number of files
    downloaded, for a quick sanity check / progress readout."""
    os.makedirs(local_dir, exist_ok=True)
    count = 0
    for item in list_children(service, folder_id):
        local_path = os.path.join(local_dir, item["name"])
        if item["mimeType"] == FOLDER_MIME:
            count += download_folder_tree(service, item["id"], local_path, _depth + 1)
        else:
            request = service.files().get_media(fileId=item["id"])
            with io.FileIO(local_path, "wb") as fh:
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            count += 1
    return count


def find_or_create_folder(service, name: str, parent_id: str) -> str:
    """Returns the ID of an existing same-named subfolder under
    parent_id, or creates one if it doesn't exist yet — so rebuilding a
    report a second time updates the same Drive folder instead of
    piling up duplicates."""
    resp = service.files().list(
        q=f"'{parent_id}' in parents and name = '{name}' and "
          f"mimeType = '{FOLDER_MIME}' and trashed = false",
        fields="files(id, name)",
    ).execute()
    existing = resp.get("files", [])
    if existing:
        return existing[0]["id"]
    folder = service.files().create(
        body={"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]},
        fields="id",
    ).execute()
    return folder["id"]


def upload_folder_tree(service, local_dir: str, parent_folder_id: str) -> int:
    """Mirrors a local directory tree UP into Drive under
    parent_folder_id, recreating the same folder structure. Existing
    same-named files are updated in place rather than duplicated, so
    re-running a build after a correction cleanly replaces the old
    version instead of leaving stale copies behind."""
    count = 0
    for entry in sorted(os.listdir(local_dir)):
        local_path = os.path.join(local_dir, entry)
        if os.path.isdir(local_path):
            sub_id = find_or_create_folder(service, entry, parent_folder_id)
            count += upload_folder_tree(service, local_path, sub_id)
        else:
            media = MediaFileUpload(local_path, resumable=True)
            existing = service.files().list(
                q=f"'{parent_folder_id}' in parents and name = '{entry}' and trashed = false",
                fields="files(id)",
            ).execute().get("files", [])
            if existing:
                service.files().update(fileId=existing[0]["id"], media_body=media).execute()
            else:
                service.files().create(
                    body={"name": entry, "parents": [parent_folder_id]}, media_body=media,
                ).execute()
            count += 1
    return count
