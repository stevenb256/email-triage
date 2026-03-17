#!/usr/bin/env python3
"""
Microsoft Graph OAuth token manager with auto-refresh.

First run: opens browser for interactive login.
After that: silently refreshes using cached refresh token.
Writes the current access token to .env as GRAPH_TOKEN=...
"""

import json
import os
import sys
import time
import threading

try:
    import msal
except ImportError:
    print("Installing msal...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "msal", "-q"])
    import msal

TOKEN_CACHE_FILE = os.path.join(os.path.dirname(__file__), ".graph_token_cache.json")
ENV_FILE = os.path.join(os.path.dirname(__file__), ".env")

# Microsoft Office public client ID (already admin-consented in most tenants)
# This is the "Microsoft Office" first-party app
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
AUTHORITY = "https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47"

# Scopes - using .default to get all pre-consented Graph permissions
# Or specify individual scopes:
SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Chat.Read",
    "https://graph.microsoft.com/ChannelMessage.Read.All",
    "https://graph.microsoft.com/Team.ReadBasic.All",
    "https://graph.microsoft.com/Channel.ReadBasic.All",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/Calendars.ReadWrite",
    "https://graph.microsoft.com/People.Read",
]


def _load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def _save_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def _update_env(token):
    """Write/update GRAPH_TOKEN in .env file."""
    lines = []
    found = False
    if os.path.exists(ENV_FILE):
        with open(ENV_FILE, "r") as f:
            for line in f:
                if line.startswith("GRAPH_TOKEN="):
                    lines.append(f"GRAPH_TOKEN={token}\n")
                    found = True
                else:
                    lines.append(line)
    if not found:
        lines.append(f"GRAPH_TOKEN={token}\n")
    with open(ENV_FILE, "w") as f:
        f.writelines(lines)


def get_token(silent_only=False):
    """Get an access token, refreshing silently if possible."""
    cache = _load_cache()
    app = msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=cache
    )

    # Try silent token acquisition first (uses refresh token)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    if silent_only:
        return None

    # Device code flow — works in CLI, no redirect URI needed
    flow = app.initiate_device_flow(SCOPES)
    if "user_code" not in flow:
        print(f"Failed to create device flow: {json.dumps(flow, indent=2)}")
        return None
    print(flow["message"])  # Prints "Go to https://microsoft.com/devicelogin and enter code XXXXXXX"
    sys.stdout.flush()
    # Give the user up to 5 minutes to complete the login
    flow["expires_at"] = time.time() + 300
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        _save_cache(cache)
        print(f"Logged in as: {result.get('id_token_claims', {}).get('preferred_username', 'unknown')}")
        return result["access_token"]
    else:
        print(f"Auth failed: {result.get('error_description', result.get('error', 'unknown'))}")
        return None


def refresh_loop(interval_minutes=60):
    """Background loop that refreshes the token and updates .env."""
    while True:
        token = get_token(silent_only=True)
        if token:
            _update_env(token)
            print(f"[{time.strftime('%H:%M:%S')}] Token refreshed and written to .env")
        else:
            print(f"[{time.strftime('%H:%M:%S')}] Silent refresh failed — interactive login needed")
            token = get_token(silent_only=False)
            if token:
                _update_env(token)
        time.sleep(interval_minutes * 60)


def start_background_refresh(interval_minutes=60):
    """Start token refresh in a background thread (call from app.py)."""
    # Do initial token fetch
    token = get_token()
    if token:
        _update_env(token)
    t = threading.Thread(target=refresh_loop, args=(interval_minutes,), daemon=True)
    t.start()
    return token


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Microsoft Graph token manager")
    parser.add_argument("--daemon", action="store_true", help="Run as daemon, refreshing every 50 minutes")
    parser.add_argument("--interval", type=int, default=50, help="Refresh interval in minutes (default: 50)")
    parser.add_argument("--print", action="store_true", dest="print_token", help="Print the token to stdout")
    args = parser.parse_args()

    token = get_token()
    if not token:
        print("Failed to get token")
        sys.exit(1)

    _update_env(token)
    print(f"Token written to {ENV_FILE}")

    if args.print_token:
        print(f"\n{token}")

    if args.daemon:
        print(f"Running as daemon, refreshing every {args.interval} minutes...")
        print("Press Ctrl+C to stop")
        try:
            refresh_loop(args.interval)
        except KeyboardInterrupt:
            print("\nStopped")
