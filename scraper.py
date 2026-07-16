"""Headed Playwright workflow for collecting current-month GST challans."""

from __future__ import annotations

import asyncio
import json
import re
from datetime import datetime
from pathlib import Path

from playwright.async_api import TimeoutError as PlaywrightTimeout
from playwright.async_api import async_playwright


CREDENTIALS_FILE = Path(__file__).with_name("credentials.json")
LOGIN_URL = "https://services.gst.gov.in/services/login"
CHALLAN_HISTORY_URL = "https://payment.gst.gov.in/payment/auth/challanhistory"
LOGOUT_URL = "https://services.gst.gov.in/services/logout"


def load_credentials() -> list[dict]:
    if not CREDENTIALS_FILE.exists():
        return []
    return json.loads(CREDENTIALS_FILE.read_text(encoding="utf-8"))


def push(state: dict, event: dict) -> None:
    state["progress_events"].append(event)


async def _fill_first(page, selectors: list[str], value: str) -> bool:
    for selector in selectors:
        try:
            await page.locator(selector).first.fill(value, timeout=8_000)
            return True
        except PlaywrightTimeout:
            continue
    return False


async def _wait_for_login(page, state: dict) -> bool:
    last_error = ""
    for _ in range(90):
        current_url = page.url.lower()
        if any(fragment in current_url for fragment in ("/auth/", "fowelcome", "dashboard")):
            return True
        for pattern in (
            r"invalid captcha",
            r"captcha.*(incorrect|invalid|wrong)",
            r"invalid username or password",
        ):
            locator = page.get_by_text(re.compile(pattern, re.I))
            try:
                if await locator.first.is_visible(timeout=250):
                    text = (await locator.first.inner_text()).strip()
                    if text and text != last_error:
                        push(state, {"type": "status", "message": f"⚠ {text}"})
                        last_error = text
                    break
            except PlaywrightTimeout:
                pass
        await asyncio.sleep(2)
    return False


def _current_month_amount(challans) -> int | float | None:
    now = datetime.now()
    for challan in challans:
        if challan.get("status") != "S":
            continue
        try:
            created = datetime.strptime(challan["chln_cre_dt"], "%d/%m/%Y %H:%M:%S")
        except (KeyError, TypeError, ValueError):
            continue
        if created.month == now.month and created.year == now.year:
            return challan.get("total_amt")
    return None


async def run_all(state: dict) -> None:
    """Attempt every configured account; failures become result events, never crashes."""
    credentials = load_credentials()
    total = len(credentials)
    if not credentials:
        push(state, {"type": "all_done", "written": 0, "skipped": 0})
        state["scraping_done"] = True
        return

    try:
        async with async_playwright() as pw:
            browser = await pw.chromium.launch(headless=False, args=["--start-maximized"])
            context = await browser.new_context(
                viewport=None,
                user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) "
                            "Chrome/133.0.0.0 Safari/537.36"),
            )
            page = await context.new_page()
            for index, account in enumerate(credentials, start=1):
                gstin = account["gstin"]
                state_name = account["state"]
                push(state, {"type": "account_start", "index": index, "total": total,
                             "gstin": gstin, "state": state_name})
                amount = None
                error = None
                try:
                    await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=30_000)
                    username_ok = await _fill_first(page, [
                        'input[ng-model*="username"]', 'input[id="loginid"]', 'input[name="username"]',
                    ], account["username"])
                    password_ok = await _fill_first(page, ['input[type="password"]'], account["password"])
                    if not username_ok or not password_ok:
                        push(state, {"type": "status", "message": "⚠ Could not auto-fill all credentials; complete them manually."})
                    push(state, {"type": "status", "message": "👉 Type the captcha in the Chrome window and click Login"})
                    if not await _wait_for_login(page, state):
                        error = "Login timed out"
                    else:
                        await page.goto(CHALLAN_HISTORY_URL, wait_until="networkidle", timeout=30_000)
                        push(state, {"type": "status", "message": "✓ Login successful — fetching challan data..."})
                        response_text = await page.evaluate(
                            """async (gstin) => {
                                const response = await fetch(
                                  'https://payment.gst.gov.in/payment/auth/challan/getlist?gstin=' + encodeURIComponent(gstin)
                                );
                                return await response.text();
                            }""", gstin)
                        amount = _current_month_amount(json.loads(response_text))
                        if amount is None:
                            error = "No PAID challan found for the current month"
                        elif gstin not in state["icici_rows"]:
                            error = "GSTIN is not present in the uploaded workbook"
                            amount = None
                        else:
                            state["workbook"]["GST Utilization Entry"].cell(
                                row=state["icici_rows"][gstin], column=6
                            ).value = amount
                except Exception as exc:  # Portal changes and transient errors must not stop later accounts.
                    error = str(exc) or "Unexpected portal error"
                    push(state, {"type": "status", "message": f"⚠ {error}"})

                found = amount is not None
                state["results"][gstin] = amount
                push(state, {"type": "account_done", "index": index, "gstin": gstin,
                             "state": state_name, "amount": amount, "found": found,
                             "error": error})
                if index < total:
                    try:
                        await page.goto(LOGOUT_URL, wait_until="domcontentloaded", timeout=15_000)
                    except Exception:
                        try:
                            await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=15_000)
                        except Exception:
                            pass
            await context.close()
            await browser.close()
    except Exception as exc:
        # If Playwright itself cannot launch, report each unattempted account visibly.
        push(state, {"type": "status", "message": f"⚠ Browser could not start: {exc}"})
        for index, account in enumerate(credentials, start=1):
            if account["gstin"] not in state["results"]:
                state["results"][account["gstin"]] = None
                push(state, {"type": "account_done", "index": index, "gstin": account["gstin"],
                             "state": account["state"], "amount": None, "found": False,
                             "error": "Browser could not start"})
    finally:
        written = sum(value is not None for value in state["results"].values())
        skipped = sum(value is None for value in state["results"].values())
        push(state, {"type": "all_done", "written": written, "skipped": skipped,
                     "skipped_gstins": [gstin for gstin, value in state["results"].items() if value is None]})
        state["scraping_done"] = True
