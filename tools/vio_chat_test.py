"""Send one harmless test message to a VIO agent and print the answer (no spec content)."""
import os
import sys
import time

import requests

BASE = os.environ.get("VIO_BASE_URL", "https://vio.automotive-wan.com:446")
AGENT = sys.argv[1] if len(sys.argv) > 1 else "Compare Expert"
token = os.environ.get("VIO_API_TOKEN") or os.environ.get("API_TOKEN")
if not token:
    sys.exit("No token: set VIO_API_TOKEN first.")
try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

body = {"model": AGENT, "messages": [
    {"role": "user", "content": "Translate into English: 警告音は条件成立後0.5秒以内に吹鳴すること。"}]}
t0 = time.time()
r = requests.post(f"{BASE}/chat/completions", json=body,
                  headers={"Authorization": f"Bearer {token}"}, timeout=120)
print("HTTP", r.status_code, f"({time.time() - t0:.1f} s)")
if r.ok:
    data = r.json()
    print("Answer:", data["choices"][0]["message"]["content"])
    print("Usage:", data.get("usage"))
else:
    print(r.text[:300])
