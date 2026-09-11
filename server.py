from __future__ import annotations
import json, os, threading
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from apscheduler.schedulers.background import BackgroundScheduler
from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from dotenv import load_dotenv

from collector import full_refresh, status_refresh

HERE=Path(__file__).resolve().parent
load_dotenv(HERE/".env")
TZ=ZoneInfo(os.getenv("TIMEZONE","Europe/London"))
app=FastAPI(title="TPN Dedicated Day Dashboard")
lock=threading.Lock()
status={"running":False,"mode":None,"stage":"Idle","last_attempt":None,"last_success":None,"last_error":None}

def run_job(mode):
    if not lock.acquire(blocking=False):
        return {"ok":False,"message":"Refresh already running"}
    status.update(running=True,mode=mode,stage="Starting",last_attempt=datetime.now(TZ).isoformat(timespec="seconds"))
    try:
        def cb(name):
            status["stage"]=name
            print(f"[server] {mode}: {name}",flush=True)
        result=full_refresh(cb) if mode=="full" else status_refresh(cb)
        status.update(last_success=datetime.now(TZ).isoformat(timespec="seconds"),last_error=None,stage="Completed")
        return {"ok":True,"mode":mode,"result":result}
    except Exception as e:
        status.update(last_error=f"{type(e).__name__}: {e}",stage="Failed")
        print("[server] ERROR",status["last_error"],flush=True)
        raise
    finally:
        status["running"]=False
        lock.release()

def start_background(mode):
    if status["running"]: return False
    def worker():
        try: run_job(mode)
        except Exception: pass
    threading.Thread(target=worker,daemon=True).start()
    return True

@app.get("/")
def home(): return FileResponse(HERE/"dashboard.html",media_type="text/html")

@app.get("/dashboard.html")
def dashboard(): return FileResponse(HERE/"dashboard.html",media_type="text/html")

@app.get("/dedicated-day-data.json")
def data():
    return FileResponse(HERE/"dedicated-day-data.json",media_type="application/json",
                        headers={"Cache-Control":"no-store, max-age=0"})

@app.get("/health")
def health(): return JSONResponse({"service":"tpn-dedicated-day","status":status})

@app.post("/refresh/{mode}")
def refresh(mode:str,x_refresh_token:str|None=Header(default=None)):
    expected=os.getenv("REFRESH_TOKEN")
    if not expected or x_refresh_token!=expected: raise HTTPException(401,"Invalid refresh token")
    if mode not in {"full","status"}: raise HTTPException(400,"Mode must be full or status")
    if not start_background(mode): return JSONResponse({"ok":False,"message":"Refresh already running"},status_code=409)
    return {"ok":True,"message":f"{mode} refresh started"}

@app.get("/admin")
def admin():
    token_required=bool(os.getenv("REFRESH_TOKEN"))
    return HTMLResponse(f"""<!doctype html><html><body style="font-family:Segoe UI,Arial;padding:30px;max-width:850px">
    <h1>TPN Dedicated Day — Render status</h1>
    <pre style="background:#f4f7fb;padding:16px;border-radius:8px">{json.dumps(status,indent=2)}</pre>
    <p>Scheduled full refresh: 08:00 Europe/London.</p>
    <p>Scheduled status refreshes: 10:00, 12:00, 14:00, 16:00, 18:00 Europe/London.</p>
    <p>Dashboard polls its JSON every 5 minutes.</p>
    <p>Manual API refresh requires the X-Refresh-Token header: {token_required}.</p>
    </body></html>""")

scheduler=BackgroundScheduler(timezone=TZ)
scheduler.add_job(lambda:start_background("full"),"cron",hour=8,minute=0,id="full_0800",replace_existing=True)
for h in (10,12,14,16,18):
    scheduler.add_job(lambda:start_background("status"),"cron",hour=h,minute=0,id=f"status_{h:02d}00",replace_existing=True)
scheduler.start()

@app.on_event("startup")
def startup_refresh():
    if os.getenv("AUTO_REFRESH_ON_START","true").lower()!="true": return
    p=HERE/"dedicated-day-data.json"
    stale=True
    if p.exists():
        try:
            payload=json.loads(p.read_text(encoding="utf-8"))
            stamp=payload.get("generated_at")
            if stamp:
                dt=datetime.fromisoformat(stamp)
                if dt.tzinfo is None: dt=dt.replace(tzinfo=TZ)
                stale=dt.astimezone(TZ).date()!=datetime.now(TZ).date()
        except Exception: pass
    if stale: start_background("full")
