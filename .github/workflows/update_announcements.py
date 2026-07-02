#!/usr/bin/env python3
"""
update_announcements.py

Runs every 15 minutes via GitHub Actions (.github/workflows/update-announcements.yml).
Commit this file to: .github/workflows/update_announcements.py  (that exact path —
the workflow calls `python .github/workflows/update_announcements.py`)

WHAT IT DOES
------------
1. Reads the published-CSV version of the "Announcements Data" Google Sheet.
2. For each row, downloads any attachments (Google Drive files) it hasn't
   already downloaded, saving them into images/.
3. Rebuilds announcements.json at the repo root in the exact shape index.html
   expects.

FOOLPROOFING PHILOSOPHY
------------------------
- A bad ROW (bad date, weird text, etc.) never kills the run — it's skipped
  and logged, everything else still processes.
- A bad ATTACHMENT (can't fetch metadata, download fails) never kills the
  row — it's dropped from that announcement's attachment list and logged.
- An UNRECOGNIZED file type is NEVER silently dropped. Images and PDFs get
  rendered inline; everything else (Word docs, Excel files, anything) gets
  a generic "download this file" card. Nothing an owner sends ever
  vanishes without a trace in the log.
- Only true showstoppers — can't reach the sheet, bad/missing credentials,
  missing required environment variables — fail the whole job loudly (so
  the GitHub email alert fires), but with a plain-English reason instead
  of a raw traceback.

DATA CONTRACT — Google Sheet "Announcements Data"
--------------------------------------------------
Columns (header row, exact names): ID, Title, DateSent, Audience, Body, Attachments

  ID           Any unique string (GUID, row number, etc). Required. Rows
               missing an ID are skipped.

  Title        Plain text. Required.

  DateSent     Preferred format: ISO 8601, e.g. 2026-06-25T09:00:00
               (In Power Automate: formatDateTime(utcNow(), 'yyyy-MM-ddTHH:mm:ss'))
               A handful of common fallback formats are also accepted
               (see DATE_FORMATS below), but ISO is the only one guaranteed
               to sort/parse correctly on the front end.

  Audience     Comma-separated tags, e.g. "All Residents, Building II Residents"
               Blank is allowed (renders with no tags).

  Body         PLAIN TEXT, not HTML. Separate paragraphs with a blank line;
               single line breaks become <br>. This script escapes the text
               and wraps it in <p> tags itself. Do NOT put raw HTML in this
               cell — it will be escaped and shown as literal text.

  Attachments  Comma-separated GOOGLE DRIVE FILE IDS (not filenames, not
               URLs). This is the "Id" output of the Power Automate Google
               Drive "Create file" action. Blank means no attachments.
               Any file type is accepted — images and PDFs render inline,
               everything else (docx, doc, xlsx, etc.) shows as a
               downloadable file card.

ENV VARS (set by the workflow)
-------------------------------
  SHEET_CSV_URL          Published CSV URL for the Announcements Data sheet
  DRIVE_FOLDER_ID         Google Drive "Announcement Attachments" folder ID
  REPO_OWNER, REPO_NAME   Used to build raw.githubusercontent.com URLs
  (GITHUB_TOKEN is passed but not used here — the workflow's own git steps
   handle the commit/push.)

The service account key is written by the workflow to /tmp/service_account.json.
"""

import csv
import html
import io
import json
import os
import re
import sys
import traceback
from datetime import datetime

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ── Config ────────────────────────────────────────────────────────────────

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# This file lives at .github/workflows/update_announcements.py, so the repo
# root is two levels up. Fall back to cwd if that doesn't look right (e.g.
# when GitHub Actions checks out to a differently named directory).
if not os.path.isdir(os.path.join(REPO_ROOT, ".git")):
    REPO_ROOT = os.getcwd()

IMAGES_DIR = os.path.join(REPO_ROOT, "images")
JSON_PATH = os.path.join(REPO_ROOT, "announcements.json")
SERVICE_ACCOUNT_FILE = "/tmp/service_account.json"

DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

# Types that render INLINE on the site (image tag or embedded PDF viewer).
# Anything not in this dict still gets downloaded and shown — just as a
# generic downloadable file card instead of an inline preview.
INLINE_MIME_TYPES = {
    "image/jpeg": (".jpg", "image"),
    "image/png": (".png", "image"),
    "image/gif": (".gif", "image"),
    "image/webp": (".webp", "image"),
    "application/pdf": (".pdf", "pdf"),
}

# Known "file card" types — just used to pick a clean extension when we
# recognize the mime type. Anything NOT listed here still works fine; we
# just fall back to whatever extension the original filename had.
FILE_MIME_TYPES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-excel": ".xls",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
    "text/plain": ".txt",
}

DATE_FORMATS = [
    "%Y-%m-%dT%H:%M:%S",     # ISO, preferred
    "%Y-%m-%d %H:%M:%S",
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %I:%M:%S %p",
    "%m/%d/%Y",
    "%Y-%m-%d",
]

REQUIRED_ENV_VARS = ["SHEET_CSV_URL", "REPO_OWNER", "REPO_NAME"]


# ── Helpers ──────────────────────────────────────────────────────────────

def log(msg):
    print(msg, flush=True)


def fail(msg):
    """A showstopper — print a clear reason and exit non-zero so the
    GitHub Actions email alert fires, without a confusing raw traceback."""
    log(f"\nFATAL: {msg}")
    sys.exit(1)


def require_env():
    missing = [v for v in REQUIRED_ENV_VARS if not os.environ.get(v)]
    if missing:
        fail(f"Missing required environment variable(s): {', '.join(missing)}. "
             f"Check the 'env:' block in update-announcements.yml.")


def parse_date(raw_value, row_id):
    raw_value = (raw_value or "").strip()
    for fmt in DATE_FORMATS:
        try:
            dt = datetime.strptime(raw_value, fmt)
            return dt.strftime("%Y-%m-%dT%H:%M:%S")
        except ValueError:
            continue
    log(f"  ! Row {row_id}: could not parse DateSent '{raw_value}' — using it as-is")
    return raw_value  # let the front end's Date() try; worst case it sorts oddly


def text_to_html(raw_text):
    """Plain text -> safe HTML. Escape first, then turn blank-line-separated
    chunks into <p> tags and single newlines into <br>."""
    raw_text = (raw_text or "").strip()
    if not raw_text:
        return ""
    escaped = html.escape(raw_text)
    paragraphs = re.split(r"\n\s*\n", escaped)
    html_paragraphs = [
        "<p>" + p.strip().replace("\n", "<br>") + "</p>"
        for p in paragraphs
        if p.strip()
    ]
    return "".join(html_paragraphs)


def parse_audience(raw_value):
    if not raw_value:
        return []
    return [a.strip() for a in raw_value.split(",") if a.strip()]


def parse_attachment_ids(raw_value):
    if not raw_value:
        return []
    return [a.strip() for a in raw_value.split(",") if a.strip()]


def fetch_sheet_rows():
    log(f"Fetching sheet CSV from {os.environ['SHEET_CSV_URL']}")
    try:
        resp = requests.get(os.environ["SHEET_CSV_URL"], timeout=30)
        resp.raise_for_status()
    except Exception as e:
        fail(f"Could not fetch the Announcements Data sheet CSV: {e}\n"
             f"  Check that the sheet is still published to the web as CSV "
             f"(File > Share > Publish to web) and the URL hasn't changed.")

    text = resp.content.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))
    rows = list(reader)

    if reader.fieldnames:
        headers = [h.strip() for h in reader.fieldnames]
        expected = {"ID", "Title", "DateSent", "Audience", "Body", "Attachments"}
        missing_cols = expected - set(headers)
        if missing_cols:
            log(f"  ! WARNING: sheet is missing expected column(s): "
                f"{', '.join(sorted(missing_cols))}. Detected columns: {headers}")

    log(f"  {len(rows)} row(s) found in sheet")
    return rows


def get_drive_service():
    if not os.path.isfile(SERVICE_ACCOUNT_FILE):
        fail(f"Service account key file not found at {SERVICE_ACCOUNT_FILE}. "
             f"Check the GOOGLE_SERVICE_ACCOUNT_KEY secret and the "
             f"'Write service account key' workflow step.")
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=DRIVE_SCOPES
        )
        return build("drive", "v3", credentials=creds, cache_discovery=False)
    except Exception as e:
        fail(f"Could not authenticate with the Google service account: {e}\n"
             f"  Check that the GOOGLE_SERVICE_ACCOUNT_KEY secret contains valid JSON.")


def existing_attachment_filenames():
    if not os.path.isdir(IMAGES_DIR):
        return set()
    return set(os.listdir(IMAGES_DIR))


def pick_extension_and_type(mime_type, original_name):
    """Decide the extension + display type for a Drive file.
    Returns (ext, type) where type is 'image', 'pdf', or 'file'."""
    if mime_type in INLINE_MIME_TYPES:
        return INLINE_MIME_TYPES[mime_type]

    if mime_type in FILE_MIME_TYPES:
        return FILE_MIME_TYPES[mime_type], "file"

    # Unrecognized mime type — fall back to whatever extension the
    # original filename had, so it's still openable once downloaded.
    original_ext = os.path.splitext(original_name or "")[1]
    return (original_ext if original_ext else ".file"), "file"


def download_attachment(drive_service, file_id, already_have):
    """Fetches metadata + (if needed) downloads a Drive file into images/.
    Returns {"filename", "type", "name"} or None if it couldn't be used."""
    try:
        meta = drive_service.files().get(fileId=file_id, fields="id, name, mimeType").execute()
    except Exception as e:
        log(f"    ! Could not fetch metadata for Drive file {file_id}: {e}")
        return None

    original_name = meta.get("name", file_id)
    mime_type = meta.get("mimeType", "")
    ext, file_type = pick_extension_and_type(mime_type, original_name)
    filename = f"{file_id}{ext}"

    if filename in already_have:
        return {"filename": filename, "type": file_type, "name": original_name}

    dest_path = os.path.join(IMAGES_DIR, filename)
    try:
        request = drive_service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        with open(dest_path, "wb") as f:
            f.write(buf.getvalue())
        log(f"    + Downloaded {filename} ({original_name})")
        already_have.add(filename)
        return {"filename": filename, "type": file_type, "name": original_name}
    except Exception as e:
        log(f"    ! Failed to download Drive file {file_id} ({original_name}): {e}")
        return None


# ── Main ─────────────────────────────────────────────────────────────────

def process_row(row, drive_service, already_have):
    """Turns one sheet row into an announcement dict, or None to skip it.
    Any unexpected error here is caught by the caller so one bad row can
    never take down the whole run."""
    row_id = (row.get("ID") or "").strip()
    title = (row.get("Title") or "").strip()

    if not row_id or not title:
        return None, "missing ID or Title"

    date_sent = parse_date(row.get("DateSent"), row_id)
    audience = parse_audience(row.get("Audience"))
    body_html = text_to_html(row.get("Body"))
    attachment_ids = parse_attachment_ids(row.get("Attachments"))

    attachments = []
    for file_id in attachment_ids:
        result = download_attachment(drive_service, file_id, already_have)
        if result:
            attachments.append({
                "url": f"{RAW_BASE}/images/{result['filename']}",
                "type": result["type"],
                "name": result["name"],
            })

    return {
        "id": row_id,
        "title": title,
        "dateSent": date_sent,
        "audience": audience,
        "body": body_html,
        "attachments": attachments,
    }, None


def main():
    require_env()

    global RAW_BASE
    RAW_BASE = f"https://raw.githubusercontent.com/{os.environ['REPO_OWNER']}/{os.environ['REPO_NAME']}/main"

    os.makedirs(IMAGES_DIR, exist_ok=True)

    rows = fetch_sheet_rows()
    drive_service = get_drive_service()
    already_have = existing_attachment_filenames()

    announcements = []
    skipped = 0

    for i, row in enumerate(rows, start=1):
        try:
            announcement, skip_reason = process_row(row, drive_service, already_have)
        except Exception as e:
            log(f"  ! Skipping sheet row {i}: unexpected error — {e}")
            log("    " + traceback.format_exc().replace("\n", "\n    "))
            skipped += 1
            continue

        if skip_reason:
            log(f"  ! Skipping sheet row {i}: {skip_reason}")
            skipped += 1
            continue

        announcements.append(announcement)

    try:
        with open(JSON_PATH, "w", encoding="utf-8") as f:
            json.dump({"announcements": announcements}, f, indent=2, ensure_ascii=False)
    except Exception as e:
        fail(f"Could not write {JSON_PATH}: {e}")

    log(f"\nWrote {len(announcements)} announcement(s) to {JSON_PATH} "
        f"({skipped} row(s) skipped).")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        fail(f"Unexpected error: {e}\n{traceback.format_exc()}")
