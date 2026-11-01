"""本機儀表板伺服器：讓網頁可以直接新增／刪除自選股，並永遠顯示最新資料。

只接受本機連線（127.0.0.1），其他電腦連不進來。
開啟方式：瀏覽器打開 http://127.0.0.1:8765/
開機自動啟動由 setup_schedule.ps1 設定。
"""
import json
import re
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import build_dashboard as B
import fetch_after_close as F
from common import DATA_DIR, log, parse_date, read_watchlist, setup_logging, write_watchlist

HOST, PORT = "127.0.0.1", 8765
LOCK = threading.Lock()          # 同一時間只處理一個新增／刪除
CODE_RE = re.compile(r"^[0-9A-Z]{4,6}$")


def latest_date() -> str | None:
    dates = B.available_dates()
    return dates[0] if dates else None


def load_data(day: str) -> dict:
    return json.loads((DATA_DIR / f"{day}.json").read_text(encoding="utf-8"))


def save_data(day: str, data: dict) -> None:
    path = DATA_DIR / f"{day}.json"
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)


def add_stock(code: str, tag: str, group: str) -> tuple[int, dict]:
    code = code.strip().upper()
    if not CODE_RE.match(code):
        return 400, {"error": "代號格式不正確"}
    with LOCK:
        watch = read_watchlist()
        if any(w["code"] == code for w in watch):
            return 409, {"error": f"{code} 已經在自選股裡了"}
        info = F.stock_directory().get(code)
        if not info:
            return 404, {"error": f"查無 {code}，請確認代號（目前支援上市、上櫃股票與 ETF）"}
        day = latest_date()
        if not day:
            return 500, {"error": "還沒有任何盤後資料"}
        data = load_data(day)
        w = {"code": code, "name": info["name"], "group": group.strip(), "tag": tag.strip() or "自選"}
        row = F.stock_row(w, parse_date(day), F.trading_days_of(data))
        if row is None:
            return 422, {"error": f"{info['name']}（{code}）在 {day} 沒有成交資料，可能暫停交易"}
        watch.append(w)
        write_watchlist(watch)
        data["stocks"] = [s for s in data.get("stocks", []) if s["code"] != code] + [row]
        save_data(day, data)
        B.build()
        log.info("新增自選股 %s %s", code, info["name"])
        return 200, {"stock": row}


def delete_stock(code: str) -> tuple[int, dict]:
    with LOCK:
        watch = read_watchlist()
        left = [w for w in watch if w["code"] != code]
        if len(left) == len(watch):
            return 404, {"error": f"{code} 不在自選股裡"}
        write_watchlist(left)
        day = latest_date()
        if day:
            data = load_data(day)
            data["stocks"] = [s for s in data.get("stocks", []) if s["code"] != code]
            save_data(day, data)
            B.build()
        log.info("刪除自選股 %s", code)
        return 200, {"ok": True}


def update_tag(code: str, tag: str) -> tuple[int, dict]:
    with LOCK:
        watch = read_watchlist()
        hit = next((w for w in watch if w["code"] == code), None)
        if not hit:
            return 404, {"error": f"{code} 不在自選股裡"}
        hit["tag"] = tag.strip() or "自選"
        write_watchlist(watch)
        day = latest_date()
        if day:
            data = load_data(day)
            for s in data.get("stocks", []):
                if s["code"] == code:
                    s["tag"] = hit["tag"]
            save_data(day, data)
            B.build()
        return 200, {"ok": True}


def lookup(q: str) -> list[dict]:
    q = q.strip().upper()
    if not q:
        return []
    held = {w["code"] for w in read_watchlist()}
    out = []
    for code, info in F.stock_directory().items():
        if code.startswith(q) or q in info["name"].upper():
            out.append({"code": code, "name": info["name"], "market": "上櫃" if info["market"] == "tpex" else "上市",
                        "held": code in held})
    # 代號完全相同 → 代號開頭相符 → 名稱相符
    out.sort(key=lambda x: (x["code"] != q, not x["code"].startswith(q), len(x["code"]), x["code"]))
    return out[:12]


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):     # 不在主控台印每個請求
        pass

    def _send(self, status: int, body: bytes, ctype: str):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, status: int, obj):
        self._send(status, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def _body(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(n) if n else b"{}"
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:      # 少數工具會用 Big5 送中文
            text = raw.decode("cp950", errors="replace")
        return json.loads(text or "{}")

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        try:
            if u.path in ("/", "/r/latest"):
                day = latest_date()
                if not day:
                    return self._send(200, "還沒有任何盤後資料，請先執行 fetch_after_close.py".encode(), "text/plain; charset=utf-8")
                return self._send(200, B.render_date(day).encode("utf-8"), "text/html; charset=utf-8")
            m = re.fullmatch(r"/r/(\d{4}-\d{2}-\d{2})", u.path)
            if m and (DATA_DIR / f"{m.group(1)}.json").exists():
                return self._send(200, B.render_date(m.group(1)).encode("utf-8"), "text/html; charset=utf-8")
            if u.path == "/api/ping":
                return self._json(200, {"ok": True})
            if u.path == "/api/latest":
                day = latest_date()
                gen = load_data(day).get("generated_at") if day else None
                return self._json(200, {"date": day, "generated_at": gen})
            if u.path == "/api/lookup":
                return self._json(200, lookup((q.get("q") or [""])[0]))
            if u.path == "/favicon.ico":
                return self._send(204, b"", "image/x-icon")
            self._json(404, {"error": "找不到頁面"})
        except Exception as e:
            log.exception("處理 %s 時出錯", self.path)
            self._json(500, {"error": str(e)})

    def do_POST(self):
        u = urlparse(self.path)
        try:
            body = self._body()
            if u.path == "/api/watchlist":
                return self._json(*add_stock(str(body.get("code", "")), str(body.get("tag", "自選")), str(body.get("group", ""))))
            if u.path == "/api/watchlist/tag":
                return self._json(*update_tag(str(body.get("code", "")), str(body.get("tag", ""))))
            if u.path == "/api/watchlist/delete":
                return self._json(*delete_stock(str(body.get("code", ""))))
            self._json(404, {"error": "找不到功能"})
        except Exception as e:
            log.exception("處理 %s 時出錯", self.path)
            self._json(500, {"error": str(e)})


def main() -> int:
    setup_logging("server")
    try:
        srv = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError:
        log.info("伺服器已經在執行中（%s:%s）", HOST, PORT)
        return 0
    log.info("儀表板伺服器啟動：http://%s:%s/", HOST, PORT)
    srv.serve_forever()
    return 0


if __name__ == "__main__":
    sys.exit(main())
