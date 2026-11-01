"""共用工具：網路請求、快取、日期轉換、紀錄。只用 Python 內建模組，不需另外安裝套件。"""
import csv
import io
import json
import logging
import ssl
import time
import urllib.parse
import urllib.request
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
CACHE_DIR = ROOT / "cache"
REPORT_DIR = ROOT / "reports"
LOG_DIR = ROOT / "logs"
for d in (DATA_DIR, CACHE_DIR, REPORT_DIR, LOG_DIR):
    d.mkdir(exist_ok=True)

# 證交所、櫃買中心的憑證缺少 Subject Key Identifier，Python 3.13 的嚴格模式會拒絕。
# 這裡只關閉「嚴格格式」這一項，憑證鏈與網域驗證仍照常進行。
_SSL = ssl.create_default_context()
_SSL.verify_flags &= ~ssl.VERIFY_X509_STRICT
_UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) tw-stock-dashboard"}

# 證交所對短時間大量請求會暫時封鎖 IP，每次請求間隔至少這麼多秒
_MIN_GAP = {"www.twse.com.tw": 3.0, "www.tpex.org.tw": 1.5, "www.taifex.com.tw": 1.0}
_last_hit: dict[str, float] = {}

log = logging.getLogger("twstock")


def setup_logging(name: str) -> None:
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(LOG_DIR / f"{name}_{date.today():%Y-%m}.log", encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    log.handlers = [fh, sh]


def _throttle(url: str) -> None:
    host = urllib.parse.urlparse(url).hostname or ""
    gap = _MIN_GAP.get(host, 0.5)
    wait = _last_hit.get(host, 0) + gap - time.time()
    if wait > 0:
        time.sleep(wait)
    _last_hit[host] = time.time()


def fetch_bytes(url: str, form: dict | None = None, tries: int = 3) -> bytes:
    data = urllib.parse.urlencode(form).encode() if form else None
    for i in range(tries):
        _throttle(url)
        try:
            req = urllib.request.Request(url, data=data, headers=_UA)
            with urllib.request.urlopen(req, timeout=40, context=_SSL) as r:
                return r.read()
        except Exception as e:  # 網路不穩時重試
            if i == tries - 1:
                raise
            log.warning("請求失敗，%s 秒後重試：%s (%s)", 5 * (i + 1), url, e)
            time.sleep(5 * (i + 1))
    raise RuntimeError("unreachable")


def fetch_json(url: str) -> dict | list:
    return json.loads(fetch_bytes(url))


def fetch_csv(url: str, form: dict) -> list[list[str]]:
    """期交所的 CSV 是 Big5 編碼。回傳去掉表頭的資料列。"""
    raw = fetch_bytes(url, form)
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("cp950", errors="replace")
    if text.lstrip().startswith("<"):
        raise RuntimeError(f"期交所回傳錯誤頁面而非 CSV（可能是日期區間不被接受）：{url}")
    rows = [r for r in csv.reader(io.StringIO(text)) if r and any(c.strip() for c in r)]
    return [[c.strip() for c in r] for r in rows[1:]]


def cached(key: str, loader, keep: bool):
    """keep=True 代表資料已定案（過去的日期／月份），存進快取以後不再重抓。"""
    path = CACHE_DIR / f"{key}.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    value = loader()
    if keep and value is not None:
        path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    return value


def xlsx_cells(raw: bytes, sheet: int = 1) -> dict[str, str]:
    """讀 Excel（.xlsx）第一個工作表，回傳 {'B2': '值', ...}。不需要安裝 openpyxl。"""
    import xml.etree.ElementTree as ET
    import zipfile
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    z = zipfile.ZipFile(io.BytesIO(raw))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(ns + "si"):
            shared.append("".join(t.text or "" for t in si.iter(ns + "t")))
    cells = {}
    for c in ET.fromstring(z.read(f"xl/worksheets/sheet{sheet}.xml")).iter(ns + "c"):
        v = c.find(ns + "v")
        val = v.text if v is not None else "".join(t.text or "" for t in c.iter(ns + "t"))
        if c.get("t") == "s" and val:
            val = shared[int(val)]
        if val:
            cells[c.get("r")] = val
    return cells


def excel_date(serial) -> date:
    from datetime import timedelta
    return date(1899, 12, 30) + timedelta(days=int(float(serial)))


def num(s) -> float | None:
    """'1,234.5' → 1234.5；'-'、'--'、空字串 → None"""
    if s is None:
        return None
    s = str(s).replace(",", "").replace("X", "").strip()
    if s in ("", "-", "--", "---", "NULL", "除權息", "除權", "除息"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def roc_to_date(s: str) -> date:
    """'115/10/30' 或 '1151030' → date(2026, 10, 30)"""
    s = s.strip()
    if "/" in s:
        y, m, d = s.split("/")
    else:
        y, m, d = s[:-4], s[-4:-2], s[-2:]
    return date(int(y) + 1911, int(m), int(d))


def parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def read_us_watchlist() -> list[dict]:
    path = ROOT / "us_watchlist.csv"
    if not path.exists():
        return []
    with path.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    return [{"symbol": r["代號"].strip(), "name": (r.get("名稱") or r["代號"]).strip()}
            for r in rows if (r.get("代號") or "").strip() and not r["代號"].startswith("#")]


def write_watchlist(items: list[dict]) -> None:
    """寫回 watchlist.csv。加上 BOM，Excel 直接打開中文也不會亂碼。"""
    path = ROOT / "watchlist.csv"
    tmp = path.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["代號", "名稱", "族群", "分類"])
        for it in items:
            w.writerow([it["code"], it.get("name", ""), it.get("group", ""), it.get("tag", "自選")])
    tmp.replace(path)


def read_watchlist() -> list[dict]:
    path = ROOT / "watchlist.csv"
    with path.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    out = []
    for r in rows:
        code = (r.get("代號") or "").strip()
        if not code or code.startswith("#"):
            continue
        out.append({
            "code": code,
            "name": (r.get("名稱") or "").strip(),
            "group": (r.get("族群") or "").strip(),
            "tag": (r.get("分類") or "自選").strip() or "自選",
        })
    return out
