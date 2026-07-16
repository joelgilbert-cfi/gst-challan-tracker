"""Local GST challan automation web application."""

from __future__ import annotations

import asyncio
import io
import json
import re
from pathlib import Path
from urllib.parse import quote

import openpyxl
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import computation
import scraper


ROOT = Path(__file__).parent
CREDENTIALS_FILE = ROOT / "credentials.json"
GSTIN_RE = re.compile(r"^\d{2}[A-Z0-9]{10,13}$")

app = FastAPI(title="GST Challan Automation")
app.mount("/static", StaticFiles(directory=ROOT / "static"), name="static")
templates = Jinja2Templates(directory=ROOT / "templates")

state = {
    "workbook": None,
    "icici_rows": {},
    "narration_month": "",
    "states_found": 0,
    "filename": "",
    "results": {},
    "progress_events": [],
    "scraping_done": False,
    "scraping_started": False,
}


def load_credentials() -> list[dict]:
    if not CREDENTIALS_FILE.exists():
        CREDENTIALS_FILE.write_text("[]\n", encoding="utf-8")
        return []
    try:
        data = json.loads(CREDENTIALS_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise RuntimeError("credentials.json contains invalid JSON") from exc
    return data if isinstance(data, list) else []


def save_credentials(data: list[dict]) -> None:
    CREDENTIALS_FILE.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")


def validate_account(state_name: str, gstin: str, username: str, password: str) -> tuple[dict | None, str | None]:
    values = {"state": state_name.strip(), "gstin": gstin.strip().upper(),
              "username": username.strip(), "password": password}
    if not all(values.values()):
        return None, "All four fields are required."
    if not GSTIN_RE.fullmatch(values["gstin"]):
        return None, "Enter a valid GSTIN."
    return values, None


def render_edit(account: dict | None = None, *, mode: str = "add", error: str | None = None):
    return templates.TemplateResponse("partials/account_row_edit.html", {
        "request": Request, "acc": account or {"state": "", "gstin": "", "username": "", "password": ""},
        "mode": mode, "error": error,
    })


def find_account(gstin: str) -> dict:
    account = next((item for item in load_credentials() if item["gstin"] == gstin.upper()), None)
    if account is None:
        raise HTTPException(status_code=404, detail="GSTIN not found")
    return account


@app.on_event("startup")
async def initialise_credentials() -> None:
    load_credentials()


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "cred_count": len(load_credentials())})


@app.get("/accounts", response_class=HTMLResponse)
async def accounts_page(request: Request):
    return templates.TemplateResponse("accounts.html", {"request": request, "credentials": load_credentials()})


@app.get("/credentials/count")
async def credentials_count():
    return {"count": len(load_credentials())}


@app.get("/credentials/new-row-form", response_class=HTMLResponse)
async def new_account_form(request: Request):
    return templates.TemplateResponse("partials/account_row_edit.html", {
        "request": request, "acc": {"state": "", "gstin": "", "username": "", "password": ""}, "mode": "add", "error": None,
    })


@app.get("/credentials/{gstin}/edit-form", response_class=HTMLResponse)
async def edit_account_form(gstin: str, request: Request):
    return templates.TemplateResponse("partials/account_row_edit.html", {
        "request": request, "acc": find_account(gstin), "mode": "edit", "error": None,
    })


@app.get("/credentials/{gstin}/row", response_class=HTMLResponse)
async def account_row(gstin: str, request: Request):
    return templates.TemplateResponse("partials/account_row.html", {
        "request": request, "acc": find_account(gstin), "index": 0,
    })


@app.post("/credentials", response_class=HTMLResponse)
async def add_account(request: Request, state_name: str = Form(alias="state"), gstin: str = Form(),
                      username: str = Form(), password: str = Form()):
    account, error = validate_account(state_name, gstin, username, password)
    credentials = load_credentials()
    if error is None and any(item["gstin"] == account["gstin"] for item in credentials):
        error = "An account with this GSTIN already exists."
    if error:
        return templates.TemplateResponse("partials/account_row_edit.html", {
            "request": request, "acc": account or {"state": state_name, "gstin": gstin, "username": username, "password": password},
            "mode": "add", "error": error,
        })
    credentials.append(account)
    save_credentials(credentials)
    return templates.TemplateResponse(
        "partials/account_row.html", {"request": request, "acc": account, "index": len(credentials)},
        headers={"HX-Trigger": "account-added"},
    )


@app.put("/credentials/{original_gstin}", response_class=HTMLResponse)
async def update_account(original_gstin: str, request: Request, state_name: str = Form(alias="state"),
                         gstin: str = Form(), username: str = Form(), password: str = Form()):
    credentials = load_credentials()
    original_index = next((i for i, item in enumerate(credentials) if item["gstin"] == original_gstin.upper()), None)
    if original_index is None:
        raise HTTPException(status_code=404, detail="GSTIN not found")
    account, error = validate_account(state_name, gstin, username, password)
    if error is None and any(i != original_index and item["gstin"] == account["gstin"] for i, item in enumerate(credentials)):
        error = "An account with this GSTIN already exists."
    if error:
        return templates.TemplateResponse("partials/account_row_edit.html", {
            "request": request, "acc": account or {"state": state_name, "gstin": gstin, "username": username, "password": password},
            "mode": "edit", "error": error,
        })
    credentials[original_index] = account
    save_credentials(credentials)
    return templates.TemplateResponse("partials/account_row.html", {"request": request, "acc": account, "index": original_index + 1})


@app.delete("/credentials/{gstin}", response_class=HTMLResponse)
async def delete_account(gstin: str):
    credentials = load_credentials()
    filtered = [item for item in credentials if item["gstin"] != gstin.upper()]
    if len(filtered) == len(credentials):
        raise HTTPException(status_code=404, detail="GSTIN not found")
    save_credentials(filtered)
    return HTMLResponse("", headers={"HX-Trigger": "account-deleted"})


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        return JSONResponse({"ok": False, "error": "Please upload an Excel (.xlsx) file"}, status_code=400)
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(await file.read()))
        result = computation.run(workbook)
    except ValueError as exc:
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=400)
    except Exception as exc:
        return JSONResponse({"ok": False, "error": f"Could not read Excel file: {exc}"}, status_code=400)
    state.update({
        "workbook": workbook, "icici_rows": result["icici_rows"], "narration_month": result["narration_month"],
        "states_found": result["states_found"], "filename": Path(file.filename).name,
        "results": {}, "progress_events": [], "scraping_done": False, "scraping_started": False,
    })
    return {"ok": True, "states_found": result["states_found"], "narration_month": result["narration_month"],
            "filename": Path(file.filename).name, "cred_count": len(load_credentials())}


@app.post("/start-scraping")
async def start_scraping():
    if state["workbook"] is None:
        return JSONResponse({"ok": False, "error": "Upload an Excel file first"}, status_code=400)
    if state["scraping_started"]:
        return JSONResponse({"ok": False, "error": "Scraping is already running"}, status_code=409)
    if not load_credentials():
        return JSONResponse({"ok": False, "error": "No accounts configured"}, status_code=400)
    state["scraping_started"] = True
    asyncio.create_task(scraper.run_all(state))
    return {"ok": True}


def format_sse_event(event: dict) -> str:
    event_type = event["type"]
    if event_type == "status":
        html = f'<span id="status-message">{event["message"]}</span>'
    elif event_type == "account_start":
        html = templates.get_template("partials/current_account.html").render(**event)
    elif event_type == "account_done":
        html = templates.get_template("partials/progress_row.html").render(**event)
    elif event_type == "all_done":
        html = templates.get_template("partials/completion.html").render(**event)
    else:
        return ""
    # SSE data must be line-prefixed. Template fragments do not contain user-controlled multiline content.
    data = "\n".join(f"data: {line}" for line in html.splitlines() if line.strip())
    return f"event: {event_type}\n{data}\n\n"


@app.get("/progress-stream")
async def progress_stream():
    async def generate():
        sent = 0
        while True:
            while sent < len(state["progress_events"]):
                yield format_sse_event(state["progress_events"][sent])
                sent += 1
            if state["scraping_done"]:
                break
            await asyncio.sleep(0.3)
    return StreamingResponse(generate(), media_type="text/event-stream", headers={
        "Cache-Control": "no-cache", "X-Accel-Buffering": "no",
    })


@app.get("/download")
async def download():
    if state["workbook"] is None:
        raise HTTPException(status_code=400, detail="No file available — please upload again")
    if not state["scraping_done"]:
        raise HTTPException(status_code=400, detail="Scraping not yet complete")
    buffer = io.BytesIO()
    state["workbook"].save(buffer)
    buffer.seek(0)
    safe_name = Path(state["filename"] or "gst_utilization.xlsx").stem + "_completed.xlsx"
    return StreamingResponse(buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(safe_name)}"})
