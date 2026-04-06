#!/usr/bin/env python3
"""
SSE MCP tools/list tester for supergateway.
Usage: python test_mcp_tools.py <server-url>
       python test_mcp_tools.py https://webapp-outlook.azurewebsites.net
"""

import sys
import json
import urllib.request
import urllib.error
import threading
import time

BASE_URL = sys.argv[1].rstrip("/") if len(sys.argv) > 1 else "https://webapp-outlook.azurewebsites.net"

message_endpoint = None
sse_ready = threading.Event()

def listen_sse():
    global message_endpoint
    req = urllib.request.Request(f"{BASE_URL}/sse")
    try:
        with urllib.request.urlopen(req, timeout=30) as res:
            last_event = None
            for raw_line in res:
                line = raw_line.decode("utf-8").strip()
                if line.startswith("event:"):
                    last_event = line[6:].strip()
                elif line.startswith("data:"):
                    data = line[5:].strip()
                    if last_event == "endpoint":
                        # data is the path e.g. /message?sessionId=abc123
                        message_endpoint = BASE_URL + data
                        print(f"   📡 Message endpoint: {message_endpoint}")
                        sse_ready.set()
                        return
                    last_event = None
    except Exception as e:
        print(f"SSE error: {e}")
        sse_ready.set()

def post(url, payload):
    data = json.dumps(payload).encode()
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=10) as res:
        body = res.read()
        return json.loads(body) if body else {}

def run():
    print(f"\n🔌 Target: {BASE_URL}\n")

    # Start SSE listener in background
    print("1️⃣  Opening SSE connection...")
    t = threading.Thread(target=listen_sse, daemon=True)
    t.start()

    # Wait for endpoint
    if not sse_ready.wait(timeout=10):
        print("❌ Timed out waiting for SSE endpoint")
        return

    if not message_endpoint:
        print("❌ No message endpoint received")
        return

    print("   ✅ SSE connected\n")

    # Step 2: initialize
    print("2️⃣  Sending initialize...")
    resp = post(message_endpoint, {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "clientInfo": {"name": "test-client", "version": "1.0.0"},
            "capabilities": {}
        }
    })
    print(f"   ✅ Response: {resp}\n")

    # Step 3: initialized notification
    print("3️⃣  Sending initialized notification...")
    post(message_endpoint, {
        "jsonrpc": "2.0",
        "method": "notifications/initialized"
    })
    print("   ✅ Done\n")

    # Step 4: tools/list
    print("4️⃣  Requesting tools/list...")
    resp = post(message_endpoint, {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/list",
        "params": {}
    })

    tools = resp.get("result", {}).get("tools", [])
    if tools:
        print(f"   ✅ {len(tools)} tools found:\n")
        for t in tools:
            print(f"      • {t['name']}: {t.get('description', '')[:60]}")
    else:
        print("   ❌ No tools returned. Response:", json.dumps(resp, indent=2))

if __name__ == "__main__":
    try:
        run()
    except urllib.error.URLError as e:
        print(f"\n❌ Connection error: {e.reason}")
    except Exception as e:
        print(f"\n❌ Error: {e}")