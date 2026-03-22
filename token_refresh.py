"""
token_refresh.py — Auto-refresh Outlook MCP token.

Strategy:
  1. Try to extract token from existing Edge browser localStorage (read-only)
  2. If Edge not available, try the app's own browser profile
  3. Submit token to all running McpOutlookLocal processes
  4. Store token in app DB for direct API calls (profile photos etc.)

The user must have Outlook Web open (or recently visited) in Edge for
strategy 1 to work. If not, `initial_login()` opens a visible browser
for the user to log in once.
"""
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time

from db import meta_get, meta_set

_APP_DIR = os.path.dirname(__file__)
_APP_BROWSER_PROFILE = os.path.join(_APP_DIR, ".browser-profile")
_OWA_URL = "https://sdf.outlook.cloud.microsoft"
_REFRESH_LOCK = threading.Lock()
_last_refresh = 0.0


# ── Browser profile paths ────────────────────────────────────────────────────

_EDGE_PROFILE = os.path.expanduser(
    "~/Library/Application Support/Microsoft Edge/Default"
)
_EDGE_LS_PATH = os.path.join(_EDGE_PROFILE, "Local Storage", "leveldb")


def _find_mcp_ports() -> list[int]:
    """Find all listening McpOutlookLocal ports via lsof."""
    try:
        out = subprocess.check_output(
            ["lsof", "-iTCP", "-sTCP:LISTEN", "-P", "-n"],
            text=True, timeout=5,
        )
    except Exception:
        return []
    ports = []
    for line in out.splitlines():
        if "McpOutloo" not in line:
            continue
        m = re.search(r":(\d+)\s+\(LISTEN\)", line)
        if m:
            ports.append(int(m.group(1)))
    return ports


# ── Token extraction ─────────────────────────────────────────────────────────

def _extract_token_via_playwright(profile_dir: str, headless: bool = True) -> str | None:
    """Open OWA in a Playwright persistent context (Edge) and extract the access token."""
    from playwright.sync_api import sync_playwright

    lock_file = os.path.join(profile_dir, "SingletonLock")
    if os.path.exists(lock_file):
        try:
            os.remove(lock_file)
        except OSError:
            pass

    edge_path = "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
    with sync_playwright() as pw:
        browser = pw.chromium.launch_persistent_context(
            profile_dir,
            headless=headless,
            executable_path=edge_path if os.path.exists(edge_path) else None,
            args=["--disable-blink-features=AutomationControlled"],
        )
        try:
            page = browser.new_page()
            page.goto(_OWA_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(5000)
            token = page.evaluate(_TOKEN_JS)
            page.close()
            return token
        finally:
            browser.close()


_TOKEN_JS = """() => {
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key.includes('accesstoken') && key.includes('outlook.office.com/')) {
            try { return JSON.parse(localStorage.getItem(key)).secret; }
            catch(e) {}
        }
    }
    return null;
}"""


def _extract_token_from_edge() -> str | None:
    """
    Try to extract the Outlook token from Edge's localStorage by opening
    a temporary copy of the Edge profile (so we don't lock the running Edge).
    """
    if not os.path.isdir(_EDGE_LS_PATH):
        return None

    # Copy Edge profile to a temp dir (only Local Storage + Cookies needed)
    tmpdir = tempfile.mkdtemp(prefix="edge-token-")
    try:
        default_dir = os.path.join(tmpdir, "Default")
        os.makedirs(default_dir, exist_ok=True)
        # Copy Local Storage
        shutil.copytree(
            os.path.join(_EDGE_PROFILE, "Local Storage"),
            os.path.join(default_dir, "Local Storage"),
        )
        # Copy cookies for SSO
        for f in ("Cookies", "Cookies-journal"):
            src = os.path.join(_EDGE_PROFILE, f)
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(default_dir, f))

        token = _extract_token_via_playwright(tmpdir, headless=True)
        return token
    except Exception as e:
        print(f"  [token-refresh] Edge extraction failed: {e}")
        return None
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ── Submit to MCP processes ──────────────────────────────────────────────────

def _submit_to_mcp_ports(token: str, ports: list[int]) -> list[int]:
    """Submit token to MCP token-entry pages. Returns list of ports that accepted."""
    from playwright.sync_api import sync_playwright

    submitted = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        try:
            for port in ports:
                url = f"http://127.0.0.1:{port}"
                page = browser.new_page()
                try:
                    page.goto(url, timeout=5000)
                    page.wait_for_timeout(500)
                    body_text = page.evaluate("() => document.body.innerText") or ""
                    if "Token Saved" in body_text or "saved" in body_text.lower():
                        print(f"  [token-refresh] Port {port}: already has token")
                        continue
                    inp = page.locator('input[type="text"], textarea, [placeholder*="token"], [placeholder*="Bearer"]')
                    if inp.count() > 0:
                        inp.fill(token)
                        btn = page.locator('button:has-text("Save token")')
                        if btn.count() > 0:
                            btn.click()
                            page.wait_for_timeout(2000)
                            result = page.evaluate("() => document.body.innerText") or ""
                            if "Token Saved" in result or "saved" in result.lower():
                                submitted.append(port)
                                print(f"  [token-refresh] Port {port}: token saved ✓")
                            else:
                                print(f"  [token-refresh] Port {port}: result unclear — {result[:80]}")
                        else:
                            print(f"  [token-refresh] Port {port}: no save button")
                    else:
                        print(f"  [token-refresh] Port {port}: no token input (may not need refresh)")
                except Exception as e:
                    print(f"  [token-refresh] Port {port}: error — {e}")
                finally:
                    page.close()
        finally:
            browser.close()
    return submitted


# ── Public API ───────────────────────────────────────────────────────────────

def refresh_token(force: bool = False) -> dict:
    """
    Main refresh flow. Returns {"ok": True, ...} on success.
    Skips if last refresh was < 5 min ago (unless force=True).
    """
    global _last_refresh
    if not force and (time.time() - _last_refresh) < 300:
        return {"ok": True, "skipped": True, "reason": "refreshed recently"}

    if not _REFRESH_LOCK.acquire(blocking=False):
        return {"ok": True, "skipped": True, "reason": "refresh already in progress"}

    try:
        return _do_refresh()
    finally:
        _REFRESH_LOCK.release()


def _do_refresh() -> dict:
    global _last_refresh

    ports = _find_mcp_ports()
    if not ports:
        return {"ok": False, "error": "No McpOutlookLocal processes found"}
    print(f"  [token-refresh] Found MCP ports: {ports}")

    # Strategy 1: Extract from Edge's localStorage
    token = _extract_token_from_edge()
    if token:
        print(f"  [token-refresh] Got token from Edge ({len(token)} chars)")
    else:
        # Strategy 2: Use app's own browser profile (needs prior initial_login)
        if os.path.isdir(_APP_BROWSER_PROFILE):
            try:
                token = _extract_token_via_playwright(_APP_BROWSER_PROFILE, headless=True)
                if token:
                    print(f"  [token-refresh] Got token from app profile ({len(token)} chars)")
            except Exception as e:
                print(f"  [token-refresh] App profile extraction failed: {e}")

    if not token:
        return {
            "ok": False,
            "error": "Could not extract token. Open Outlook Web in Edge, or run initial login.",
        }

    # Submit to MCP processes
    submitted = _submit_to_mcp_ports(token, ports)

    # Store in app DB
    meta_set("outlook_token", token)
    print(f"  [token-refresh] Token stored in DB")

    _last_refresh = time.time()
    return {"ok": True, "ports_submitted": submitted, "ports_found": ports}


def initial_login() -> dict:
    """
    Launch a visible Edge browser for the user to log in to OWA.
    Uses Edge (not Chromium) so corporate SSO / device certificates work.
    Establishes SSO cookies in the app's browser profile.
    """
    from playwright.sync_api import sync_playwright

    lock_file = os.path.join(_APP_BROWSER_PROFILE, "SingletonLock")
    if os.path.exists(lock_file):
        try:
            os.remove(lock_file)
        except OSError:
            pass

    print("  [token-refresh] Opening Edge for OWA login…")

    with sync_playwright() as pw:
        browser = pw.chromium.launch_persistent_context(
            _APP_BROWSER_PROFILE,
            headless=False,
            executable_path="/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = browser.new_page()
        page.goto(_OWA_URL, wait_until="domcontentloaded", timeout=60000)
        # Poll for token (up to 3 min)
        token = None
        for _ in range(180):
            try:
                page.wait_for_timeout(1000)
                token = page.evaluate(_TOKEN_JS)
            except Exception:
                break
            if token:
                print(f"  [token-refresh] Login successful! Token found ({len(token)} chars)")
                break
        try:
            page.close()
        except Exception:
            pass

        if not token:
            try:
                browser.close()
            except Exception:
                pass
            return {"ok": False, "error": "Browser closed before login completed."}

        # Submit token to MCP ports while we still have the browser context
        ports = _find_mcp_ports()
        submitted = []
        if ports:
            print(f"  [token-refresh] Submitting to MCP ports: {ports}")
            for port in ports:
                url = f"http://127.0.0.1:{port}"
                tp = browser.new_page()
                try:
                    tp.goto(url, timeout=5000)
                    tp.wait_for_timeout(500)
                    body_text = tp.evaluate("() => document.body.innerText") or ""
                    if "Token Saved" in body_text:
                        print(f"  [token-refresh] Port {port}: already has token")
                        continue
                    inp = tp.locator('input[type="text"], textarea, [placeholder*="token"], [placeholder*="Bearer"]')
                    if inp.count() > 0:
                        inp.fill(token)
                        btn = tp.locator('button:has-text("Save token")')
                        if btn.count() > 0:
                            btn.click()
                            tp.wait_for_timeout(2000)
                            result_text = tp.evaluate("() => document.body.innerText") or ""
                            if "Token Saved" in result_text or "saved" in result_text.lower():
                                submitted.append(port)
                                print(f"  [token-refresh] Port {port}: token saved ✓")
                except Exception as e:
                    print(f"  [token-refresh] Port {port}: error — {e}")
                finally:
                    tp.close()

        try:
            browser.close()
        except Exception:
            pass

        # Store token in DB
        global _last_refresh
        meta_set("outlook_token", token)
        _last_refresh = time.time()
        print(f"  [token-refresh] Token stored in DB")
        return {"ok": True, "ports_submitted": submitted, "ports_found": ports, "source": "initial_login"}


def needs_refresh() -> bool:
    """Quick check: is the stored token expired or missing?"""
    token = meta_get("outlook_token", "")
    if not token:
        return True
    try:
        import base64
        payload = token.split(".")[1]
        payload += "=" * (4 - len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload))
        exp = claims.get("exp", 0)
        return time.time() > (exp - 300)
    except Exception:
        return (time.time() - _last_refresh) > 3000
