#!/usr/bin/env python3
"""
Streamable HTTP MCP tools/list tester.
Usage: python test_mcp_tools.py <server-url>
       python test_mcp_tools.py https://webapp-outlook.azurewebsites.net
"""

import sys
import json
import urllib.request
import urllib.error
BASE_URL = sys.argv[1].rstrip("/") if len(sys.argv) > 1 else "https://webapp-outlook.azurewebsites.net"
MCP_ENDPOINT = f"{BASE_URL}/mcp"

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
    print(f"1️⃣  Using MCP endpoint: {MCP_ENDPOINT}\n")

    # Step 2: initialize
    print("2️⃣  Sending initialize...")
    resp = post(MCP_ENDPOINT, {
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
    post(MCP_ENDPOINT, {
        "jsonrpc": "2.0",
        "method": "notifications/initialized"
    })
    print("   ✅ Done\n")

    # Step 4: tools/list
    print("4️⃣  Requesting tools/list...")
    resp = post(MCP_ENDPOINT, {
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