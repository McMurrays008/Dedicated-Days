from __future__ import annotations

import os
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from collector import manual_morning_import, status_refresh_from_file

HERE = Path(__file__).resolve().parent
SYNC_DIR = HERE / "onedrive-sync"
SYNC_DIR.mkdir(exist_ok=True)


def _stage(cb, name: str):
    print(f"[onedrive] {name}", flush=True)
    if cb:
        cb(name)


def _required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"{name} is not configured")
    return value


def _download_url(shared_url: str) -> str:
    """Turn an Anyone-with-the-link OneDrive/SharePoint URL into a direct download URL."""
    parts = urllib.parse.urlsplit(shared_url.strip())
    query = urllib.parse.parse_qsl(parts.query, keep_blank_values=True)
    query = [(k, v) for k, v in query if k.lower() != "download"]
    query.append(("download", "1"))
    return urllib.parse.urlunsplit(
        (parts.scheme, parts.netloc, parts.path, urllib.parse.urlencode(query), parts.fragment)
    )


def _download(shared_url: str, target: Path, expected: str) -> Path:
    url = _download_url(shared_url)
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "*/*",
        },
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        body = resp.read()
        content_type = (resp.headers.get("Content-Type") or "").lower()

    if not body:
        raise RuntimeError(f"{expected} download was empty")

    if expected == "xlsx":
        if not body.startswith(b"PK"):
            sample = body[:200].decode("utf-8", errors="replace")
            raise RuntimeError(
                "Morning OneDrive link did not return an Excel workbook. "
                f"Content-Type={content_type!r}; response starts {sample!r}"
            )
    elif expected == "csv":
        # A CSV should be plain text, not the OneDrive/SharePoint HTML preview page.
        head = body[:500].decode("utf-8", errors="replace").lower()
        if "<html" in head or "<!doctype" in head:
            raise RuntimeError(
                "Status OneDrive link returned a web page instead of the CSV file. "
                "Check that the link is set to Anyone with the link / Can view."
            )

    target.write_bytes(body)
    return target


def sync_from_onedrive(stage_callback=None):
    """Refresh the dashboard directly from two Anyone-with-the-link OneDrive files."""
    tz = ZoneInfo(os.getenv("TIMEZONE", "Europe/London"))
    today = datetime.now(tz).date()

    morning_url = _required("ONEDRIVE_MORNING_URL")
    status_url = _required("ONEDRIVE_STATUS_URL")

    _stage(stage_callback, "Downloading TPN Dedicated Day Check from OneDrive")
    morning_path = _download(
        morning_url,
        SYNC_DIR / "TPN Dedicated Day Check.xlsx",
        "xlsx",
    )

    _stage(stage_callback, f"Importing morning population for {today.strftime('%d/%m/%Y')}")
    morning_result = manual_morning_import(morning_path, stage_callback)

    _stage(stage_callback, "Downloading Browse Export from OneDrive")
    status_path = _download(
        status_url,
        SYNC_DIR / "Browse Export.csv",
        "csv",
    )

    _stage(stage_callback, "Applying OneDrive statuses by Docket")
    status_result = status_refresh_from_file(status_path, stage_callback)

    return {
        "date": today.isoformat(),
        "morning": morning_result,
        "status": status_result,
        "source": "OneDrive Anyone links",
    }
