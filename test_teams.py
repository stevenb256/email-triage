#!/usr/bin/env python3
"""Test Microsoft Teams access via Graph API."""

import requests
import json

TOKEN = "YOUR_GRAPH_TOKEN_HERE"  # Replace with a valid token obtained via Graph Explorer

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
}

def test_endpoint(name, url, params=None):
    """Test an API endpoint and print results."""
    print(f"\n{'='*60}")
    print(f"Testing: {name}")
    print(f"URL: {url}")
    try:
        resp = requests.get(url, headers=HEADERS, params=params, timeout=10)
        print(f"Status: {resp.status_code}")
        if resp.status_code == 200:
            data = resp.json()
            if "value" in data:
                items = data["value"]
                print(f"Results: {len(items)} items")
                for i, item in enumerate(items[:5]):
                    if "displayName" in item:
                        print(f"  [{i+1}] {item.get('displayName', 'N/A')} (id: {item.get('id', 'N/A')[:20]}...)")
                    elif "topic" in item or "chatType" in item:
                        print(f"  [{i+1}] topic: {item.get('topic', '(no topic)')} type: {item.get('chatType', 'N/A')}")
                    elif "body" in item:
                        body = item.get("body", {}).get("content", "")[:100]
                        sender = item.get("from", {}).get("user", {}).get("displayName", "Unknown") if item.get("from") else "Unknown"
                        print(f"  [{i+1}] {sender}: {body}")
                    else:
                        print(f"  [{i+1}] {json.dumps(item, indent=2)[:200]}")
                if len(items) > 5:
                    print(f"  ... and {len(items) - 5} more")
            else:
                print(json.dumps(data, indent=2)[:500])
        else:
            error_text = resp.text[:400]
            print(f"Error: {error_text}")
    except Exception as e:
        print(f"Exception: {e}")

def decode_token_scopes():
    """Decode JWT payload to show audience and scopes."""
    import base64
    parts = TOKEN.split(".")
    # Pad base64
    payload = parts[1] + "=" * (4 - len(parts[1]) % 4)
    data = json.loads(base64.urlsafe_b64decode(payload))
    print(f"Token audience: {data.get('aud')}")
    print(f"Token expires: {data.get('exp')}")
    print(f"User: {data.get('upn')}")
    scopes = data.get("scp", "").split()
    teams_scopes = [s for s in scopes if any(kw in s for kw in ["Chat", "Channel", "Team"])]
    print(f"\nTeams-related scopes in token ({len(teams_scopes)}):")
    for s in teams_scopes:
        print(f"  - {s}")
    return data

if __name__ == "__main__":
    print("Microsoft Teams API Test (Graph-scoped token)")
    print("=" * 60)

    # Decode token
    print("\n--- Token Info ---")
    token_data = decode_token_scopes()

    GRAPH = "https://graph.microsoft.com/v1.0"
    BETA = "https://graph.microsoft.com/beta"

    # Profile sanity check
    test_endpoint("My Profile", f"{GRAPH}/me")

    # Teams
    test_endpoint("My Joined Teams", f"{GRAPH}/me/joinedTeams")

    # Chats
    test_endpoint("My Chats (recent 10)", f"{GRAPH}/me/chats",
                  params={"$top": "10", "$orderby": "lastUpdatedDateTime desc"})

    # Chats with last message preview
    test_endpoint("Chats with last message",
                  f"{GRAPH}/me/chats",
                  params={"$top": "5", "$expand": "lastMessagePreview"})

    # Beta chats (often has more fields)
    test_endpoint("Chats (beta, with members)",
                  f"{BETA}/me/chats",
                  params={"$top": "5", "$expand": "members"})

    # Fetch channels for first few teams
    print(f"\n\n--- Fetching channels for first 3 teams ---")
    resp = requests.get(f"{GRAPH}/me/joinedTeams", headers=HEADERS, timeout=10)
    if resp.status_code == 200:
        teams = resp.json().get("value", [])[:3]
        for team in teams:
            tid = team["id"]
            tname = team["displayName"]
            test_endpoint(f"Channels in '{tname}'",
                          f"{GRAPH}/teams/{tid}/channels")
            # Try to get recent messages from the General channel
            ch_resp = requests.get(f"{GRAPH}/teams/{tid}/channels", headers=HEADERS, timeout=10)
            if ch_resp.status_code == 200:
                channels = ch_resp.json().get("value", [])
                general = next((c for c in channels if c.get("displayName") == "General"), channels[0] if channels else None)
                if general:
                    test_endpoint(f"Messages in '{tname}/General'",
                                  f"{GRAPH}/teams/{tid}/channels/{general['id']}/messages",
                                  params={"$top": "3"})
