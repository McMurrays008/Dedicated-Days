from __future__ import annotations

import json
import os
import urllib.error
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


def _token() -> str:
    tenant = _required("MS_TENANT_ID")
    client_id = _required("MS_CLIENT_ID")
    client_secret = _required("MS_CLIENT_SECRET")
    url = f"https://login.microsoftonline.com/{urllib.parse.quote(tenant)}/oauth2/v2.0/token"
    body = urllib.parse.urlencode(
        {
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }
    ).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Microsoft token request failed: HTTP {e.code}: {detail[:500]}")
    token = payload.get("access_token")
    if not token:
        raise RuntimeError("Microsoft token response did not contain an access_token")
    return token


def _graph_json(url: str, token: str) -> dict:
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Microsoft Graph request failed: HTTP {e.code}: {detail[:800]}")


def _download(url: str, token: str, target: Path) -> Path:
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            target.write_bytes(resp.read())
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OneDrive download failed: HTTP {e.code}: {detail[:500]}")
    if not target.exists() or target.stat().st_size < 10:
        raise RuntimeError(f"Downloaded OneDrive file is empty: {target.name}")
    return target


def _children(token: str) -> list[dict]:
    user = _required("ONEDRIVE_USER")
    folder = os.getenv("ONEDRIVE_FOLDER_PATH", "Customer Service/Dedicated Days").strip().strip("/")
    user_q = urllib.parse.quote(user, safe="")
    folder_q = urllib.parse.quote(folder, safe="/")
    url = (
        f"https://graph.microsoft.com/v1.0/users/{user_q}/drive/root:/{folder_q}:/children"
        "?$select=id,name,lastModifiedDateTime,size,file"
    )
    items: list[dict] = []
    while url:
        payload = _graph_json(url, token)
        items.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return items


def _latest(items: list[dict], predicate, label: str) -> dict:
    matches = [x for x in items if x.get("file") is not None and predicate(str(x.get("name", "")))]
    if not matches:
        raise RuntimeError(f"No {label} file was found in the OneDrive Dedicated Days folder")
    matches.sort(key=lambda x: str(x.get("lastModifiedDateTime", "")), reverse=True)
    return matches[0]


def _download_item(item: dict, token: str, stem: str) -> Path:
    user = _required("ONEDRIVE_USER")
    user_q = urllib.parse.quote(user, safe="")
    item_id = urllib.parse.quote(str(item["id"]), safe="")
    name = str(item.get("name", ""))
    suffix = Path(name).suffix.lower() or ".dat"
    target = SYNC_DIR / f"{stem}{suffix}"
    url = f"https://graph.microsoft.com/v1.0/users/{user_q}/drive/items/{item_id}/content"
    return _download(url, token, target)


def sync_from_onedrive(stage_callback=None):
    """Load today's morning population and latest status export from OneDrive.

    The OneDrive folder is the source of truth. Each run re-imports today's
    TPN Dedicated Day Check, then applies the newest Browse Export /
    ConsignmentExport status file by Docket.
    """
    tz = ZoneInfo(os.getenv("TIMEZONE", "Europe/London"))
    today = datetime.now(tz).date()

    _stage(stage_callback, "Connecting to Microsoft OneDrive")
    token = _token()

    _stage(stage_callback, "Reading Dedicated Days folder")
    items = _children(token)

    morning = _latest(
        items,
        lambda n: n.lower().startswith("tpn dedicated day check") and n.lower().endswith(".xlsx"),
        "TPN Dedicated Day Check",
    )
    status = _latest(
        items,
        lambda n: (
            n.lower().startswith("browse export")
            or n.lower().startswith("consignmentexport")
            or n.lower().startswith("consignment export")
        )
        and Path(n).suffix.lower() in {".csv", ".xlsx"},
        "Browse Export / ConsignmentExport",
    )

    _stage(stage_callback, f"Downloading morning file: {morning['name']}")
    morning_path = _download_item(morning, token, "morning-dedicated-day-check")

    _stage(stage_callback, f"Importing morning population for {today.strftime('%d/%m/%Y')}")
    morning_result = manual_morning_import(morning_path, stage_callback)

    _stage(stage_callback, f"Downloading status file: {status['name']}")
    status_path = _download_item(status, token, "browse-export")

    _stage(stage_callback, "Applying OneDrive statuses by Docket")
    status_result = status_refresh_from_file(status_path, stage_callback)

    return {
        "date": today.isoformat(),
        "morning_file": morning.get("name"),
        "morning_modified": morning.get("lastModifiedDateTime"),
        "status_file": status.get("name"),
        "status_modified": status.get("lastModifiedDateTime"),
        "morning": morning_result,
        "status": status_result,
    }
