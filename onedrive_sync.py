from __future__ import annotations

import html
import json
import os
import re
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


def _fetch_bytes(url: str) -> tuple[bytes, str, str]:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/154 Safari/537.36",
            "Accept": "*/*",
        },
    )
    with urllib.request.urlopen(req, timeout=120) as resp:
        return (
            resp.read(),
            (resp.headers.get("Content-Type") or "").lower(),
            resp.geturl(),
        )


def _extract_download_url(page_html: str, base_url: str) -> str | None:
    """Best-effort extraction of the real file URL from a SharePoint anonymous preview page."""
    text = html.unescape(page_html)

    # SharePoint/Office pages often carry the file URL in JSON-like boot data.
    patterns = [
        r'"downloadUrl"\s*:\s*"([^"]+)"',
        r'"DownloadUrl"\s*:\s*"([^"]+)"',
        r'"downloadURL"\s*:\s*"([^"]+)"',
        r'"@microsoft\.graph\.downloadUrl"\s*:\s*"([^"]+)"',
        r'"FileGetUrl"\s*:\s*"([^"]+)"',
        r'"fileGetUrl"\s*:\s*"([^"]+)"',
        r'href=["\']([^"\']+/_layouts/(?:15/)?download\.aspx[^"\']*)["\']',
    ]
    for pattern in patterns:
        m = re.search(pattern, text, flags=re.I)
        if not m:
            continue
        raw = m.group(1)
        # Decode common JSON escaping used in Microsoft boot payloads.
        try:
            raw = json.loads(f'"{raw}"')
        except Exception:
            raw = raw.replace("\\u0026", "&").replace("\\u003d", "=").replace("\\/", "/")
        raw = html.unescape(raw)
        if raw.startswith("//"):
            raw = "https:" + raw
        return urllib.parse.urljoin(base_url, raw)

    # Last-resort search for an absolute URL containing download.aspx.
    m = re.search(r'(https?://[^"\'<> ]+/_layouts/(?:15/)?download\.aspx[^"\'<> ]*)', text, flags=re.I)
    if m:
        return html.unescape(m.group(1).replace("\\u0026", "&").replace("\\/", "/"))
    return None


def _download(shared_url: str, target: Path, expected: str) -> Path:
    # First try the normal SharePoint download switch.
    url = _download_url(shared_url)
    body, content_type, final_url = _fetch_bytes(url)

    # Some modern anonymous SharePoint links still return an HTML preview page
    # even with download=1. In that case, follow the real download URL embedded
    # in the page boot data.
    head = body[:1000].decode("utf-8", errors="replace").lower()
    if "text/html" in content_type or "<html" in head or "<!doctype" in head:
        page_html = body.decode("utf-8", errors="replace")
        candidate = _extract_download_url(page_html, final_url)
        if candidate:
            body, content_type, final_url = _fetch_bytes(candidate)

    if not body:
        raise RuntimeError(f"{expected} download was empty")

    if expected == "xlsx":
        if not body.startswith(b"PK"):
            raise RuntimeError(
                "Morning OneDrive link still returned a SharePoint web page instead of the Excel workbook. "
                "Please use the file's direct Download link rather than the normal sharing link."
            )
    elif expected == "csv":
        head = body[:1000].decode("utf-8", errors="replace").lower()
        if "text/html" in content_type or "<html" in head or "<!doctype" in head:
            raise RuntimeError(
                "Status OneDrive link still returned a SharePoint web page instead of the CSV file. "
                "Please use the file's direct Download link rather than the normal sharing link."
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
