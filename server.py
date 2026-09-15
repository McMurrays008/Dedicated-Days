from __future__ import annotations

import json
import os
import threading
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from apscheduler.schedulers.background import BackgroundScheduler
from dotenv import load_dotenv
from fastapi import FastAPI, Header, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

from collector import full_refresh, status_refresh, manual_morning_import

HERE = Path(__file__).resolve().parent
load_dotenv(HERE / ".env")
TZ = ZoneInfo(os.getenv("TIMEZONE", "Europe/London"))
app = FastAPI(title="TPN Dedicated Day Dashboard")
lock = threading.Lock()
status = {
    "running": False,
    "mode": None,
    "stage": "Idle",
    "last_attempt": None,
    "last_success": None,
    "last_error": None,
}


def run_job(mode):
    if not lock.acquire(blocking=False):
        return {"ok": False, "message": "Refresh already running"}

    status.update(
        running=True,
        mode=mode,
        stage="Starting",
        last_attempt=datetime.now(TZ).isoformat(timespec="seconds"),
    )

    try:
        def cb(name):
            status["stage"] = name
            print(f"[server] {mode}: {name}", flush=True)

        result = full_refresh(cb) if mode == "full" else status_refresh(cb)
        status.update(
            last_success=datetime.now(TZ).isoformat(timespec="seconds"),
            last_error=None,
            stage="Completed",
        )
        return {"ok": True, "mode": mode, "result": result}
    except Exception as e:
        status.update(last_error=f"{type(e).__name__}: {e}", stage="Failed")
        print("[server] ERROR", status["last_error"], flush=True)
        raise
    finally:
        status["running"] = False
        lock.release()


def start_background(mode):
    if status["running"]:
        return False

    def worker():
        try:
            run_job(mode)
        except Exception:
            pass

    threading.Thread(target=worker, daemon=True).start()
    return True


@app.get("/")
def home():
    return FileResponse(HERE / "dashboard.html", media_type="text/html")


@app.get("/dashboard.html")
def dashboard():
    return FileResponse(HERE / "dashboard.html", media_type="text/html")


@app.get("/dedicated-day-data.json")
def data():
    return FileResponse(
        HERE / "dedicated-day-data.json",
        media_type="application/json",
        headers={"Cache-Control": "no-store, max-age=0"},
    )


@app.get("/health")
def health():
    return JSONResponse(
        {
            "service": "tpn-dedicated-day",
            "status": status,
            "refresh_token_configured": bool(os.getenv("REFRESH_TOKEN")),
        }
    )


@app.post("/refresh/{mode}")
def refresh(mode: str, x_refresh_token: str | None = Header(default=None)):
    expected = os.getenv("REFRESH_TOKEN")
    if not expected:
        raise HTTPException(503, "REFRESH_TOKEN is not configured in Render")
    if x_refresh_token != expected:
        raise HTTPException(401, "Invalid refresh token")
    if mode not in {"full", "status"}:
        raise HTTPException(400, "Mode must be full or status")
    if not start_background(mode):
        return JSONResponse(
            {"ok": False, "message": "Refresh already running"}, status_code=409
        )
    return {"ok": True, "message": f"{mode} refresh started"}



@app.post("/morning-import")
async def morning_import(file: UploadFile = File(...), x_refresh_token: str | None = Header(default=None)):
    expected=os.getenv("REFRESH_TOKEN")
    if not expected:
        raise HTTPException(503,"REFRESH_TOKEN is not configured in Render")
    if x_refresh_token != expected:
        raise HTTPException(401,"Invalid refresh token")
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400,"Please upload the TPN Dedicated Day Check .xlsx file")
    target=HERE/"morning-dedicated-day-check.xlsx"
    target.write_bytes(await file.read())
    try:
        result=manual_morning_import(target)
        status.update(last_success=datetime.now(TZ).isoformat(timespec="seconds"),last_error=None,stage="Morning import completed")
        return {"ok":True,"result":result}
    except Exception as e:
        status.update(last_error=f"{type(e).__name__}: {e}",stage="Morning import failed")
        raise HTTPException(500,str(e))

@app.get("/admin")
def admin():
    configured = "true" if os.getenv("REFRESH_TOKEN") else "false"
    return HTMLResponse(
        f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>TPN Dedicated Day Admin</title>
<style>
body{{font-family:Segoe UI,Arial,sans-serif;background:#f5f8fc;color:#172033;margin:0}}
.wrap{{max-width:920px;margin:36px auto;padding:0 18px}}
.card{{background:#fff;border:1px solid #dbe5f0;border-radius:12px;padding:22px;box-shadow:0 4px 18px rgba(20,50,90,.06);margin-bottom:18px}}
h1{{margin:0 0 8px;font-size:28px}} h2{{margin-top:0;font-size:19px}}
.muted{{color:#65758b}} .row{{display:flex;gap:12px;flex-wrap:wrap;align-items:end}}
.field{{flex:1;min-width:260px}} label{{display:block;font-weight:600;margin-bottom:6px}}
input{{width:100%;box-sizing:border-box;padding:11px;border:1px solid #bfcddd;border-radius:8px;font-size:15px}}
button{{border:0;border-radius:8px;padding:11px 16px;font-weight:700;cursor:pointer;background:#1769e0;color:#fff}}
button.secondary{{background:#42566f}} button:disabled{{opacity:.5;cursor:not-allowed}}
#result{{white-space:pre-wrap;background:#f4f7fb;padding:14px;border-radius:8px;min-height:24px}}
pre{{background:#f4f7fb;padding:16px;border-radius:8px;overflow:auto}}
.ok{{color:#12733c;font-weight:700}} .bad{{color:#b42318;font-weight:700}}
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <h1>TPN Dedicated Day — Admin</h1>
    <p class="muted">Run a manual refresh without Render Shell access.</p>
    <p id="tokenState"></p>
    <div class="row">
      <div class="field">
        <label for="token">Refresh token</label>
        <input id="token" type="password" autocomplete="off" placeholder="Paste your REFRESH_TOKEN from Render">
      </div>
      <div class="field">
        <label for="morningFile">Morning TPN Dedicated Day Check (.xlsx)</label>
        <input id="morningFile" type="file" accept=".xlsx">
      </div>
      <button onclick="uploadMorning()">Import Morning Check</button>
      <button class="secondary" onclick="runRefresh('status')">Update Status Now</button>
    </div>
    <p class="muted">Import the morning Dedicated Day Check once. That fixes the day's delivery population. Status refreshes only update those Dockets; they do not add or remove deliveries.</p>
    <div id="result">Ready.</div>
  </div>

  <div class="card">
    <h2>Current refresh status</h2>
    <pre id="statusBox">Loading…</pre>
  </div>

  <div class="card">
    <h2>Schedule</h2>
    <p>Morning population: <strong>manual import at about 09:30 Europe/London</strong>.</p>
    <p>Automatic Browse status refreshes: <strong>10:00, 12:00, 14:00, 16:00, 18:00 Europe/London</strong>.</p>
    <p>Dashboard checks its JSON every 5 minutes.</p>
  </div>
</div>
<script>
const tokenConfigured = {configured};
const tokenState = document.getElementById('tokenState');
tokenState.innerHTML = tokenConfigured
  ? '<span class="ok">REFRESH_TOKEN is configured in Render.</span>'
  : '<span class="bad">REFRESH_TOKEN is NOT configured in Render. Add it under Environment before using the buttons.</span>';

async function uploadMorning() {{
  const result = document.getElementById('result');
  const token = document.getElementById('token').value.trim();
  const file = document.getElementById('morningFile').files[0];
  if (!token) {{ result.textContent='Paste the REFRESH_TOKEN first.'; return; }}
  if (!file) {{ result.textContent='Choose the morning TPN Dedicated Day Check .xlsx file first.'; return; }}
  const form=new FormData();
  form.append('file',file);
  result.textContent='Importing morning Dedicated Day Check…';
  try {{
    const r=await fetch('/morning-import',{{method:'POST',headers:{{'X-Refresh-Token':token}},body:form}});
    result.textContent='HTTP '+r.status+'\n'+await r.text();
    await loadStatus();
  }} catch(e) {{
    result.textContent='Import failed: '+e;
  }}
}}

async function runRefresh(mode) {{
  const result = document.getElementById('result');
  const token = document.getElementById('token').value.trim();
  if (!tokenConfigured) {{ result.textContent = 'REFRESH_TOKEN is not configured in Render.'; return; }}
  if (!token) {{ result.textContent = 'Paste the REFRESH_TOKEN first.'; return; }}
  result.textContent = 'Starting ' + mode + ' refresh…';
  try {{
    const r = await fetch('/refresh/' + mode, {{
      method: 'POST',
      headers: {{'X-Refresh-Token': token}}
    }});
    const text = await r.text();
    result.textContent = 'HTTP ' + r.status + '\n' + text;
    await loadStatus();
  }} catch (e) {{
    result.textContent = 'Request failed: ' + e;
  }}
}}

async function loadStatus() {{
  try {{
    const r = await fetch('/health?ts=' + Date.now(), {{cache:'no-store'}});
    const j = await r.json();
    document.getElementById('statusBox').textContent = JSON.stringify(j, null, 2);
  }} catch (e) {{
    document.getElementById('statusBox').textContent = 'Could not load status: ' + e;
  }}
}}
loadStatus();
setInterval(loadStatus, 5000);
</script>
</body>
</html>"""
    )


scheduler = BackgroundScheduler(timezone=TZ)
for h in (10, 12, 14, 16, 18):
    scheduler.add_job(
        lambda: start_background("status"),
        "cron",
        hour=h,
        minute=0,
        id=f"status_{h:02d}00",
        replace_existing=True,
    )
scheduler.start()


@app.on_event("startup")
def startup_refresh():
    # Morning population is intentionally manual. Scheduled Browse status
    # refreshes begin only after that day's file has been imported.
    return
