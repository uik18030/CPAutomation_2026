import os 
import sys

import requests

BASE = os.environ.get("VIO_BASE_URL", "https://vio.automotive-wan.com:446")
token = os.environ.get("VIO_API_TOKEN") or os.environ.get("API_TOKEN")
if not token:
    sys.exitt("No token: set VIO_API_TOKEN (or API_TOKEN) first.")

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

r = requests.get(f"{BASE}/models", headers={"Authorization": f"Bearer {token}"}, timeout=20)
print("HTTP", r.status_code)
if r.ok:
    for m in r.json().get("data", []):
        print(" -", m.get("id"))
else:
    print(r.text[:300])