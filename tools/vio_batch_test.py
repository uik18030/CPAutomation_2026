"""Send 10 harmless test sentences in ONE request and time it (no spec content)."""
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

lines = ["警告音は条件成立後0.5秒以内に吹鳴すること。", "IG-OFF判定成立でエンディングモードへ遷移すること。",
         "メーター照明はイルミ信号に従い調光すること。", "燃料残量警告はFuel Level信号が10%以下で点灯すること。",
         "表示言語の切替はセットアップ画面から行うこと。", "シフト位置表示はShift Position信号に従い表示すること。",
         "外気温表示は-40℃から60℃の範囲とすること。", "時計表示は12時間表示と24時間表示を切り替え可能とすること。",
         "診断要求を受信した場合は、診断仕様書に従い応答すること。", "開始要求はIG-ON判定成立時とする。"]
prompt = ("Translate each numbered line into English. Use 'shall' for すること. Keep codes, numbers and units "
          "unchanged. Answer only with the same numbering, one line each, nothing else.\n"
          + "\n".join(f"{i}. {t}" for i, t in enumerate(lines, 1)))
t0 = time.time()
r = requests.post(f"{BASE}/chat/completions", json={"model": AGENT, "messages": [{"role": "user", "content": prompt}]},
                  headers={"Authorization": f"Bearer {token}"}, timeout=300)
print("HTTP", r.status_code, f"({time.time() - t0:.1f} s)")
if r.ok:
    data = r.json()
    print(data["choices"][0]["message"]["content"].strip())
    print("Usage:", data.get("usage"))
else:
    print(r.text[:300])