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

from playwright.sync_api import sync_playwright
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


def _download_with_browser(shared_url: str, target: Path) -> Path:
    """Download an anonymously shared SharePoint/OneDrive file the same way a browser user does."""
    with sync_playwright() as p:
        browser=p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox","--disable-setuid-sandbox","--disable-dev-shm-usage",
                "--disable-gpu","--disable-extensions","--renderer-process-limit=1",
                "--disable-background-networking","--disable-background-timer-throttling",
                "--disable-backgrounding-occluded-windows","--disable-breakpad",
                "--disable-component-update","--disable-default-apps",
                "--disable-notifications","--disable-sync","--no-first-run",
                "--no-default-browser-check","--mute-audio"
            ],
        )
        context=browser.new_context(
            accept_downloads=True,
            viewport={"width":1100,"height":760},
            service_workers="block",
        )
        page=context.new_page()
        page.set_default_timeout(20000)

        download_holder={"download":None}

        def got_download(d):
            download_holder["download"]=d

        page.on("download",got_download)
        try:
            try:
                page.goto(shared_url,wait_until="domcontentloaded",timeout=45000)
            except Exception as e:
                # A direct download URL can cause navigation to abort because a
                # download starts immediately; keep going if a download event fired.
                if download_holder["download"] is None:
                    print(f"[onedrive] Browser navigation note: {e}",flush=True)

            for _ in range(20):
                if download_holder["download"] is not None:
                    download_holder["download"].save_as(target)
                    return target
                page.wait_for_timeout(500)

            def wait_for_download(seconds=12):
                loops=max(1,int(seconds*2))
                for _ in range(loops):
                    if download_holder["download"] is not None:
                        download_holder["download"].save_as(target)
                        return True
                    page.wait_for_timeout(500)
                return False

            # Once the anonymous-share page has established its cookies/session,
            # try its current URL with download=1 inside the same browser context.
            try:
                current=page.url
                parts=urllib.parse.urlsplit(current)
                q=urllib.parse.parse_qsl(parts.query,keep_blank_values=True)
                q=[(k,v) for k,v in q if k.lower()!="download"]
                q.append(("download","1"))
                forced=urllib.parse.urlunsplit((parts.scheme,parts.netloc,parts.path,urllib.parse.urlencode(q),parts.fragment))
                try:
                    page.goto(forced,wait_until="domcontentloaded",timeout=30000)
                except Exception:
                    pass
                if wait_for_download(8):
                    return target
            except Exception:
                pass

            def click_download():
                candidates=[
                    page.get_by_role("button",name=re.compile(r"download",re.I)),
                    page.get_by_role("link",name=re.compile(r"download",re.I)),
                    page.get_by_role("menuitem",name=re.compile(r"download",re.I)),
                    page.get_by_text(re.compile(r"^\s*download\s*$",re.I)),
                    page.locator("a[download]"),
                    page.locator("[aria-label*='Download' i]"),
                    page.locator("[title*='Download' i]"),
                    page.locator("[data-automationid*='download' i]"),
                    page.locator("[data-testid*='download' i]"),
                ]
                for loc in candidates:
                    try:
                        for i in range(min(loc.count(),30)):
                            el=loc.nth(i)
                            if not el.is_visible():
                                continue
                            try:
                                el.click(timeout=8000)
                            except Exception:
                                el.evaluate("(el)=>el.click()")
                            return True
                    except Exception:
                        pass
                return False

            clicked=click_download()

            # Microsoft sometimes hides Download under a More / ellipsis menu.
            if not clicked:
                more_candidates=[
                    page.get_by_role("button",name=re.compile(r"(more|more options|see more)",re.I)),
                    page.locator("[aria-label*='More' i]"),
                    page.locator("[title*='More' i]"),
                    page.locator("button:has-text('...')"),
                    page.locator("[data-automationid*='more' i]"),
                ]
                for loc in more_candidates:
                    try:
                        for i in range(min(loc.count(),20)):
                            el=loc.nth(i)
                            if not el.is_visible():
                                continue
                            try:
                                el.click(timeout=5000)
                            except Exception:
                                el.evaluate("(el)=>el.click()")
                            page.wait_for_timeout(750)
                            if click_download():
                                clicked=True
                                break
                    except Exception:
                        pass
                    if clicked:
                        break

            if clicked and wait_for_download(45):
                return target

            # Put useful diagnostics in Render logs if Microsoft changes the UI.
            visible=[]
            for sel in ("button","a","[role='menuitem']","[aria-label]","[title]"):
                try:
                    q=page.locator(sel)
                    for i in range(min(q.count(),80)):
                        el=q.nth(i)
                        if not el.is_visible():
                            continue
                        visible.append({
                            "tag":sel,
                            "text":" ".join((el.inner_text() or "").split())[:100],
                            "aria":(el.get_attribute("aria-label") or "")[:100],
                            "title":(el.get_attribute("title") or "")[:100],
                        })
                except Exception:
                    pass
            print(f"[onedrive] SHAREPOINT UI DIAGNOSTIC url={page.url!r} visible={visible[:120]}",flush=True)
            raise RuntimeError(
                "Could not download the anonymous SharePoint file after trying the page, "
                "download=1 in-session, and the More/Download menus."
            )
        finally:
            try: page.remove_listener("download",got_download)
            except Exception: pass
            try: page.close()
            except Exception: pass
            try: context.close()
            except Exception: pass
            try: browser.close()
            except Exception: pass



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
            print("[onedrive] HTTP route returned SharePoint HTML; trying browser download",flush=True)
            return _download_with_browser(shared_url,target)
    elif expected == "csv":
        head = body[:1000].decode("utf-8", errors="replace").lower()
        if "text/html" in content_type or "<html" in head or "<!doctype" in head:
            print("[onedrive] HTTP route returned SharePoint HTML; trying browser download",flush=True)
            return _download_with_browser(shared_url,target)

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
