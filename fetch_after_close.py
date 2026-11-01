"""每日盤後資料抓取。

用法：
    python fetch_after_close.py                 抓今天（非交易日會自動略過）
    python fetch_after_close.py --date 2026-10-30   補抓指定日期
    python fetch_after_close.py --wait          資料還沒公布時，每 15 分鐘重試，直到 18:30

結果存到 data/YYYY-MM-DD.json，給 build_dashboard.py 產生儀表板。
"""
import argparse
import json
import sys
import time
import urllib.parse
from datetime import date, datetime, timedelta, timezone

from common import (DATA_DIR, cached, excel_date, fetch_bytes, fetch_csv, fetch_json, log, num,
                    parse_date, read_us_watchlist, read_watchlist, roc_to_date, setup_logging,
                    xlsx_cells, CACHE_DIR)

TWSE = "https://www.twse.com.tw/rwd/zh"
TPEX = "https://www.tpex.org.tw/www/zh-tw"
TAIFEX = "https://www.taifex.com.tw/cht/3"

HISTORY_MONTHS = 4      # 個股往回抓幾個月（算 60 日均線需要約 3 個月）
INST_DAYS = 20          # 法人買賣超往回看幾個交易日（算連買連賣，最多算到 20 天）


# ───────────────────────── 交易日判斷 ─────────────────────────

def holidays(year: int) -> set[date]:
    def load():
        j = fetch_json(f"{TWSE}/holidaySchedule/holidaySchedule?response=json&queryYear={year - 1911}")
        out = []
        for d, name, *_ in j.get("data", []):
            # 「開始交易日」「最後交易日」是有開盤的日子，不算休市
            if "交易日" in name and "無交易" not in name:
                continue
            out.append(d)
        return out
    return {parse_date(d) for d in cached(f"holidays_{year}", load, keep=True)}


def is_trading_day(d: date) -> bool:
    return d.weekday() < 5 and d not in holidays(d.year)


def months_back(d: date, n: int) -> list[date]:
    out, y, m = [], d.year, d.month
    for _ in range(n):
        out.append(date(y, m, 1))
        m -= 1
        if m == 0:
            y, m = y - 1, 12
    return list(reversed(out))


# ───────────────────────── 大盤 ─────────────────────────

def taiex_series(target: date) -> list[dict]:
    rows = []
    for m in months_back(target, HISTORY_MONTHS):
        def load(m=m):
            j = fetch_json(f"{TWSE}/afterTrading/FMTQIK?date={m:%Y%m%d}&response=json")
            return j.get("data", []) if j.get("stat") == "OK" else []
        done = (m.year, m.month) < (target.year, target.month)
        for r in cached(f"twse_FMTQIK_{m:%Y%m}", load, keep=done):
            rows.append({
                "d": roc_to_date(r[0]).isoformat(),
                "close": num(r[4]),
                "change": num(r[5]),
                "turnover": round(num(r[2]) / 1e8, 2),   # 億元
            })
    return [r for r in rows if r["d"] <= target.isoformat()]


# ───────────────────────── 三大法人（大盤） ─────────────────────────

def bfi82u(d: date) -> dict | None:
    def load():
        j = fetch_json(f"{TWSE}/fund/BFI82U?type=day&dayDate={d:%Y%m%d}&response=json")
        if j.get("stat") != "OK":
            return None
        v = {r[0].strip(): num(r[3]) / 1e8 for r in j["data"]}
        get = lambda *keys: sum(val for k, val in v.items() if any(key in k for key in keys))
        return {
            "foreign": round(get("外資及陸資", "外資自營商"), 2),
            "trust": round(get("投信"), 2),
            "dealer_self": round(get("自營商(自行"), 2),
            "dealer_hedge": round(get("自營商(避險"), 2),
            "total": round(v.get("合計", 0), 2),
        }
    return cached(f"twse_BFI82U_{d:%Y%m%d}", load, keep=True)


# ───────────────────────── 三大法人（個股） ─────────────────────────

def t86(d: date) -> dict | None:
    """上市個股法人買賣超，單位：張。{代號: [外資, 投信, 自營商, 合計]}"""
    def load():
        j = fetch_json(f"{TWSE}/fund/T86?date={d:%Y%m%d}&selectType=ALLBUT0999&response=json")
        if j.get("stat") != "OK":
            return None
        f = j["fields"]
        idx = lambda name: f.index(name)
        i_f1 = idx("外陸資買賣超股數(不含外資自營商)")
        i_f2 = idx("外資自營商買賣超股數")
        i_t = idx("投信買賣超股數")
        i_d = idx("自營商買賣超股數")
        i_all = idx("三大法人買賣超股數")
        out = {}
        for r in j["data"]:
            g = lambda i: (num(r[i]) or 0) / 1000
            out[r[0].strip()] = [round(g(i_f1) + g(i_f2)), round(g(i_t)), round(g(i_d)), round(g(i_all))]
        return out
    return cached(f"twse_T86_{d:%Y%m%d}", load, keep=True)


def tpex_insti(d: date) -> dict | None:
    """上櫃個股法人買賣超，單位：張。"""
    def load():
        j = fetch_json(f"{TPEX}/insti/dailyTrade?type=Daily&sect=EW&date={d:%Y}%2F{d:%m}%2F{d:%d}&response=json")
        t = (j.get("tables") or [{}])[0]
        if not t.get("data") or roc_to_date(t.get("date", "0/1/1")) != d:
            return None
        out = {}
        for r in t["data"]:
            g = lambda i: (num(r[i]) or 0) / 1000
            # 欄位：外資合計=10、投信=13、自營商合計=22、三大法人合計=23
            out[r[0].strip()] = [round(g(10)), round(g(13)), round(g(22)), round(g(23))]
        return out
    return cached(f"tpex_insti_{d:%Y%m%d}", load, keep=True)


# ───────────────────────── 個股日成交 ─────────────────────────

def stock_month(code: str, market: str, m: date, target: date) -> list[dict]:
    def load():
        if market == "twse":
            j = fetch_json(f"{TWSE}/afterTrading/STOCK_DAY?date={m:%Y%m%d}&stockNo={code}&response=json")
            data = j.get("data", []) if j.get("stat") == "OK" else []
            vol_div = 1000          # 股 → 張
        else:
            j = fetch_json(f"{TPEX}/afterTrading/tradingStock?code={code}&date={m:%Y}%2F{m:%m}%2F01&response=json")
            data = (j.get("tables") or [{}])[0].get("data", [])
            vol_div = 1             # 已經是張
        rows = []
        for r in data:
            close = num(r[6])
            if close is None:
                continue
            rows.append({
                "d": roc_to_date(r[0]).isoformat(),
                "open": num(r[3]), "high": num(r[4]), "low": num(r[5]), "close": close,
                "change": num(r[7]) or 0.0,
                "vol": round((num(r[1]) or 0) / vol_div),
            })
        return rows
    done = (m.year, m.month) < (target.year, target.month)
    return cached(f"{market}_day_{code}_{m:%Y%m}", load, keep=done)


def detect_market(code: str, target: date) -> str | None:
    path = CACHE_DIR / "markets.json"
    known = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    if code in known:
        return known[code]
    m = date(target.year, target.month, 1)
    found = None
    for market in ("twse", "tpex"):
        if stock_month(code, market, m, target) or stock_month(code, market, months_back(target, 2)[0], target):
            found = market
            break
    if found:
        known[code] = found
        path.write_text(json.dumps(known, ensure_ascii=False, indent=1), encoding="utf-8")
    return found


# ───────────────────────── 期貨與選擇權 ─────────────────────────

def futures_tx(d: date) -> dict | None:
    rows = fetch_csv(f"{TAIFEX}/futDataDown", {
        "down_type": "1", "commodity_id": "TX",
        "queryStartDate": f"{d:%Y/%m/%d}", "queryEndDate": f"{d:%Y/%m/%d}"})
    day = [r for r in rows if r[1] == "TX" and r[17] == "一般" and r[2].isdigit() and len(r[2]) == 6]
    if not day:
        return None
    r = min(day, key=lambda r: r[2])     # 近月合約
    return {"contract": r[2], "open": num(r[3]), "high": num(r[4]), "low": num(r[5]),
            "close": num(r[6]), "change": num(r[7]), "volume": num(r[9]),
            "settle": num(r[10]), "oi": num(r[11])}


def futures_institutions(d: date) -> list[dict]:
    rows = fetch_csv(f"{TAIFEX}/futContractsDateDown", {
        "queryStartDate": f"{d - timedelta(days=45):%Y/%m/%d}",
        "queryEndDate": f"{d:%Y/%m/%d}", "commodityId": "TXF"})
    by_day: dict[str, dict] = {}
    key = {"外資及陸資": "foreign", "投信": "trust", "自營商": "dealer"}
    for r in rows:
        who = key.get(r[2])
        if not who:
            continue
        day = r[0].replace("/", "-")
        by_day.setdefault(day, {"d": day})[who] = int(num(r[13]) or 0)
    return [by_day[k] for k in sorted(by_day)]


def pc_ratio(d: date) -> list[dict]:
    # 這個下載最多只接受約一個月的區間
    rows = fetch_csv(f"{TAIFEX}/pcRatioDown", {
        "queryStartDate": f"{d - timedelta(days=30):%Y/%m/%d}", "queryEndDate": f"{d:%Y/%m/%d}"})
    out = [{"d": r[0].replace("/", "-"), "vol": num(r[3]), "oi": num(r[6])} for r in rows]
    return sorted(out, key=lambda x: x["d"])


def options_txo(d: date, ref_price: float) -> dict | None:
    rows = fetch_csv(f"{TAIFEX}/optDataDown", {
        "down_type": "1", "commodity_id": "TXO",
        "queryStartDate": f"{d:%Y/%m/%d}", "queryEndDate": f"{d:%Y/%m/%d}"})
    # 只留「一般」時段、且在 d 之後才到期的合約（d 當天到期的已經結算，不再有意義）
    today = f"{d:%Y%m%d}"
    rows = [r for r in rows if len(r) > 20 and r[1] == "TXO" and r[17] == "一般" and r[20] > today]
    if not rows:
        return None

    def strike_table(rs) -> list[dict]:
        oi: dict[float, dict] = {}
        for r in rs:
            k = num(r[3])
            side = "c" if r[4] == "買權" else "p"
            oi.setdefault(k, {"k": k, "c": 0, "p": 0})[side] += int(num(r[11]) or 0)
        return sorted((s for s in oi.values() if s["c"] or s["p"]), key=lambda s: s["k"])

    def max_pain(table) -> float | None:
        # 結算價落在哪裡時，買方的總價值最小（也就是賣方付出最少）
        if not table:
            return None
        pain = lambda x: sum(s["c"] * max(0, x - s["k"]) + s["p"] * max(0, s["k"] - x) for s in table)
        return min(table, key=lambda s: pain(s["k"]))["k"]

    def contract_info(rs) -> dict:
        e = rs[0][20]
        return {"contract": rs[0][2], "expiry": f"{e[:4]}-{e[4:6]}-{e[6:]}", "max_pain": max_pain(strike_table(rs))}

    # 最近到期的合約（多半是週選）
    near_exp = min(r[20] for r in rows)
    nearest = contract_info([r for r in rows if r[20] == near_exp])
    # 最近的月選（到期月份是 6 碼數字）
    monthly_rows = [r for r in rows if r[2].isdigit() and len(r[2]) == 6]
    monthly = None
    if monthly_rows:
        m_exp = min(r[20] for r in monthly_rows)
        monthly = contract_info([r for r in monthly_rows if r[20] == m_exp])

    # 所有未到期合約加總。壓力＝現價以上 10% 內買權未平倉最多的履約價；
    # 支撐＝現價以下 10% 內賣權未平倉最多的履約價（只看價外，避免深價內的策略部位干擾）
    total = strike_table(rows)
    above = [s for s in total if ref_price <= s["k"] <= ref_price * 1.10] or total
    below = [s for s in total if ref_price * 0.90 <= s["k"] <= ref_price] or total
    call_wall = max(above, key=lambda s: s["c"])
    put_wall = max(below, key=lambda s: s["p"])
    return {
        "nearest": nearest, "monthly": monthly,
        "call_wall": {"k": call_wall["k"], "oi": call_wall["c"]},
        "put_wall": {"k": put_wall["k"], "oi": put_wall["p"]},
        "call_oi_total": sum(s["c"] for s in total), "put_oi_total": sum(s["p"] for s in total),
        "strikes": [s for s in total if abs(s["k"] - ref_price) <= 2500 and s["k"] % 100 == 0],
    }


# ───────────────────────── 美股、匯率（Yahoo Finance） ─────────────────────────

def us_markets(d: date) -> list[dict]:
    """抓 us_watchlist.csv 裡每個代號「台灣 d 日 16:30 時已知」的最新收盤。
    美股取美東 d-1 日（含）以前的最後一個交易日；匯率取 d 日（含）以前。"""
    out = []
    for w in read_us_watchlist():
        sym = w["symbol"]
        kind = ("fx" if sym.endswith("=X") else "yield" if sym == "^TNX"
                else "index" if sym.startswith("^") else "stock")
        try:
            q = urllib.parse.quote(sym, safe="")
            j = fetch_json(f"https://query1.finance.yahoo.com/v8/finance/chart/{q}?range=1mo&interval=1d")
            r = j["chart"]["result"][0]
            off = r["meta"].get("gmtoffset", 0)
            cutoff = d if kind == "fx" else d - timedelta(days=1)
            bars = []
            for t, c in zip(r.get("timestamp", []), r["indicators"]["quote"][0].get("close", [])):
                if c is None:
                    continue
                local = datetime.fromtimestamp(t + off, tz=timezone.utc).date()   # 交易所當地日期
                if local <= cutoff:
                    bars.append((local, c))
            if len(bars) < 2:
                raise ValueError("資料不足")
            (dt, last), (_, prev) = bars[-1], bars[-2]
            out.append({"symbol": sym, "name": w["name"], "kind": kind, "date": dt.isoformat(),
                        "close": round(last, 4), "change": round(last - prev, 4),
                        "pct": round((last - prev) / prev * 100, 2)})
        except Exception as e:
            log.warning("美股／匯率 %s 抓取失敗：%s", sym, e)
    return out


# ───────────────────────── 三大法人買賣超排行 ─────────────────────────

def quotes_all(d: date) -> dict | None:
    """上市櫃全部股票當日行情 {代號: [名稱, 收盤, 漲跌, 市場]}。"""
    def load():
        out = {}
        j = fetch_json(f"{TWSE}/afterTrading/MI_INDEX?date={d:%Y%m%d}&type=ALLBUT0999&response=json")
        t = next((t for t in j.get("tables") or [] if "收盤價" in (t.get("fields") or []) and "證券代號" in t["fields"]), None)
        if not t:
            return None
        f = t["fields"]
        i_n, i_c, i_s, i_d = f.index("證券名稱"), f.index("收盤價"), f.index("漲跌(+/-)"), f.index("漲跌價差")
        for r in t["data"]:
            close = num(r[i_c])
            if close is None:
                continue
            diff = num(r[i_d]) or 0
            out[r[0].strip()] = [r[i_n].strip(), close, -diff if "-" in r[i_s] else diff, "twse"]
        try:
            j = fetch_json(f"{TPEX}/afterTrading/otc?date={d:%Y}%2F{d:%m}%2F{d:%d}&type=EW&response=json")
            t = (j.get("tables") or [{}])[0]
            if t.get("data") and roc_to_date(t.get("date", "0/1/1")) == d:
                for r in t["data"]:
                    close = num(r[2])
                    if close is not None:
                        out.setdefault(r[0].strip(), [r[1].strip(), close, num(r[3]) or 0, "tpex"])
        except Exception as e:
            log.warning("上櫃行情抓取失敗：%s", e)
        return out
    return cached(f"quotes_all_{d:%Y%m%d}", load, keep=True)


def inst_ranking(d: date, top: int = 10) -> dict | None:
    """外資、投信、自營商、三大法人合計，各自買超／賣超前 N 名（依金額與依張數兩種排法）。"""
    nets = t86(d)
    if not nets:
        return None
    nets = {**nets, **(tpex_insti(d) or {})}
    q = quotes_all(d) or {}
    rows = []
    for code, (fo, tr, de, tot) in nets.items():
        info = q.get(code)
        if not info:
            continue
        name, close, chg, market = info
        prev = close - chg
        rows.append({"code": code, "name": name, "market": market, "close": close,
                     "pct": round(chg / prev * 100, 2) if prev else None,
                     "etf": code.startswith("00"), "v": {"foreign": fo, "trust": tr, "dealer": de, "total": tot}})
    out = {}
    for who in ("foreign", "trust", "dealer", "total"):
        def item(r):
            lots = r["v"][who]
            return {k: r[k] for k in ("code", "name", "market", "close", "pct", "etf")} | {
                "lots": lots, "amount": round(lots * 1000 * r["close"] / 1e8, 2)}
        items = [item(r) for r in rows if r["v"][who]]
        out[who] = {
            "buy": {"amount": sorted((i for i in items if i["lots"] > 0), key=lambda i: -i["amount"])[:top],
                    "lots": sorted((i for i in items if i["lots"] > 0), key=lambda i: -i["lots"])[:top]},
            "sell": {"amount": sorted((i for i in items if i["lots"] < 0), key=lambda i: i["amount"])[:top],
                     "lots": sorted((i for i in items if i["lots"] < 0), key=lambda i: i["lots"])[:top]},
        }
    return out


# ───────────────────────── 美日台融資 ─────────────────────────

def twse_margin_total(d: date) -> dict | None:
    """上市融資餘額（億元、張）。證交所信用交易資料通常晚上才公布。"""
    def load():
        j = fetch_json(f"{TWSE}/marginTrading/MI_MARGN?date={d:%Y%m%d}&selectType=MS&response=json")
        t = (j.get("tables") or [{}])[0]
        rows = {r[0]: r for r in t.get("data") or []}
        lots, amt = rows.get("融資(交易單位)"), rows.get("融資金額(仟元)")
        if j.get("stat") != "OK" or not lots or not amt:
            return None
        return {"d": d.isoformat(), "balance": round(num(amt[5]) / 1e5, 2), "prev": round(num(amt[4]) / 1e5, 2),
                "lots": num(lots[5])}
    return cached(f"twse_MARGN_MS_{d:%Y%m%d}", load, keep=True)


def twse_maintenance(d: date, balance_100m: float) -> float | None:
    """大盤融資維持率 ＝ Σ(個股融資餘額 × 收盤價) ÷ 融資金額。"""
    def load_margin():
        j = fetch_json(f"{TWSE}/marginTrading/MI_MARGN?date={d:%Y%m%d}&selectType=ALL&response=json")
        t = next((t for t in j.get("tables") or [] if "彙總" in (t.get("title") or "")), None)
        return {r[0].strip(): num(r[6]) or 0 for r in t["data"]} if t and t.get("data") else None

    def load_close():
        j = fetch_json(f"{TWSE}/afterTrading/MI_INDEX?date={d:%Y%m%d}&type=ALLBUT0999&response=json")
        t = next((t for t in j.get("tables") or [] if "收盤價" in (t.get("fields") or []) and "證券代號" in t["fields"]), None)
        if not t:
            return None
        i = t["fields"].index("收盤價")
        return {r[0].strip(): num(r[i]) for r in t["data"] if num(r[i])}

    bal = cached(f"twse_MARGN_ALL_{d:%Y%m%d}", load_margin, keep=True)
    close = cached(f"twse_CLOSE_ALL_{d:%Y%m%d}", load_close, keep=True)
    if not bal or not close or not balance_100m:
        return None
    value = sum(lots * 1000 * close[c] for c, lots in bal.items() if c in close)
    return round(value / (balance_100m * 1e8) * 100, 2)


def tw_margin(days: list[date]) -> dict | None:
    """days 由新到舊。當天還沒公布時，用最近一個已公布的交易日。"""
    hist = []
    for d in days:
        v = twse_margin_total(d)
        if v:
            hist.append(v)
    if not hist:
        return None
    latest = hist[0]
    return {
        "date": latest["d"], "balance": latest["balance"],
        "change": round(latest["balance"] - latest["prev"], 2), "lots": latest["lots"],
        "maintenance": twse_maintenance(parse_date(latest["d"]), latest["balance"]),
        "history": [{"d": h["d"], "v": h["balance"]} for h in hist[::-1]],
    }


def us_margin(d: date) -> dict | None:
    """FINRA 每月公布的美股融資餘額（Margin Debt），單位：百萬美元。約在次月第三週更新。"""
    raw = fetch_bytes_cached("finra_margin", "https://www.finra.org/sites/default/files/2021-03/margin-statistics.xlsx")
    cells = xlsx_cells(raw)
    rows = []
    r = 2
    while f"A{r}" in cells:
        ym, v = cells[f"A{r}"], num(cells.get(f"B{r}"))
        if v and ym <= f"{d:%Y-%m}":
            rows.append({"d": ym, "v": v})
        r += 1
    rows.sort(key=lambda x: x["d"])
    if len(rows) < 13:
        return None
    last, prev, year_ago = rows[-1], rows[-2], rows[-13]
    return {"date": last["d"], "balance": last["v"],
            "mom": round((last["v"] / prev["v"] - 1) * 100, 2),
            "yoy": round((last["v"] / year_ago["v"] - 1) * 100, 2),
            "history": rows[-24:]}


def jp_margin(d: date) -> dict | None:
    """日本交易所每週公布的信用取引現在高（東京＋名古屋），金額單位：百萬日圓。"""
    import re
    page = fetch_bytes("https://www.jpx.co.jp/markets/statistics-equities/margin/04.html").decode("utf-8", "replace")
    files = sorted(set(re.findall(r'href="([^"]+?/(\d{8})_mtcurrent\.xlsx)"', page)), key=lambda x: x[1])
    weeks = []
    for path, stamp in files:
        def load(path=path):
            c = xlsx_cells(fetch_bytes("https://www.jpx.co.jp" + path))
            # 第 14 列＝信用取引残高合計（株數）、第 15 列＝金額；E＝賣殘、G＝買殘、H＝買殘前週比
            return {"d": excel_date(c["B2"]).isoformat(),
                    "buy": num(c["G15"]), "buy_chg": num(c["H15"]), "sell": num(c["E15"]),
                    "ratio": round(num(c["G14"]) / num(c["E14"]), 2)}
        try:
            w = cached(f"jpx_margin_{stamp}", load, keep=True)
            if w["d"] <= d.isoformat():
                weeks.append(w)
        except Exception as e:
            log.warning("日股融資檔 %s 解析失敗：%s", path, e)
    if not weeks:
        return None
    weeks.sort(key=lambda w: w["d"])
    last = weeks[-1]
    return {"date": last["d"], "balance": last["buy"], "change": last["buy_chg"], "sell": last["sell"],
            "ratio": last["ratio"], "history": [{"d": w["d"], "v": w["buy"]} for w in weeks]}


def fetch_bytes_cached(key: str, url: str) -> bytes:
    """同一天只下載一次（存在 cache/當天日期）。"""
    path = CACHE_DIR / f"{key}_{date.today():%Y%m%d}.bin"
    if path.exists():
        return path.read_bytes()
    raw = fetch_bytes(url)
    path.write_bytes(raw)
    return raw


# ───────────────────────── 計算工具 ─────────────────────────

def streak(values: list[float]) -> int:
    """values 由新到舊。回傳 +N＝連 N 日買超，-N＝連 N 日賣超。"""
    n, sign = 0, 0
    for v in values:
        s = (v > 0) - (v < 0)
        if s == 0:
            break
        if sign == 0:
            sign = s
        if s != sign:
            break
        n += 1
    return n * sign


def ma(closes: list[float], n: int) -> float | None:
    return round(sum(closes[-n:]) / n, 2) if len(closes) >= n else None


def stock_signals(s: dict, closes: list[float]) -> list[str]:
    sig = []
    c = s["close"]
    if s["pct"] is not None:
        if s["pct"] >= 9.5:
            sig.append("漲停")
        elif s["pct"] <= -9.5:
            sig.append("跌停")
    if len(closes) >= 21 and s["ma20"]:
        prev_ma20 = sum(closes[-21:-1]) / 20
        if c > s["ma20"] and closes[-2] <= prev_ma20:
            sig.append("站上月線")
        elif c < s["ma20"] and closes[-2] >= prev_ma20:
            sig.append("跌破月線")
    if s["ma5"] and s["ma20"] and s["ma60"]:
        if c > s["ma5"] > s["ma20"] > s["ma60"]:
            sig.append("多頭排列")
        elif c < s["ma5"] < s["ma20"] < s["ma60"]:
            sig.append("空頭排列")
    if len(closes) >= 60:
        if c >= s["high60"]:
            sig.append("60日新高")
        elif c <= s["low60"]:
            sig.append("60日新低")
    if s["vol_ratio"]:
        if s["vol_ratio"] >= 2:
            sig.append("爆量")
        elif s["vol_ratio"] <= 0.5:
            sig.append("量縮")
    for key, who in (("foreign_streak", "外資"), ("trust_streak", "投信")):
        n = s.get(key) or 0
        more = "+" if abs(n) >= INST_DAYS else ""   # 達到觀察上限，實際可能更久
        if n >= 3:
            sig.append(f"{who}連{n}{more}買")
        elif n <= -3:
            sig.append(f"{who}連{-n}{more}賣")
    return sig


# ───────────────────────── 個股整理 ─────────────────────────

def inst_days(days: list[date], need_tpex: bool) -> list[dict]:
    """每個交易日的個股法人買賣超（上市＋視需要加上櫃），days 由新到舊。"""
    out = []
    for d in days:
        merged = dict(t86(d) or {})
        if need_tpex:
            merged.update(tpex_insti(d) or {})
        out.append(merged)
    return out


def stock_row(w: dict, target: date, days: list[date], inst_by_day: list[dict] | None = None,
              market: str | None = "auto") -> dict | None:
    """算出一檔自選股在 target 日的所有欄位。回傳 None 代表當天沒有成交資料。"""
    if market == "auto":
        market = detect_market(w["code"], target)
    if not market:
        log.warning("找不到 %s 的交易資料，請確認代號", w["code"])
        return {**w, "market": None, "error": "查無此代號"}
    if inst_by_day is None:
        inst_by_day = inst_days(days, need_tpex=market == "tpex")
    rows = []
    for m in months_back(target, HISTORY_MONTHS):
        rows += stock_month(w["code"], market, m, target)
    rows = [r for r in rows if r["d"] <= target.isoformat()]
    if not rows or rows[-1]["d"] != target.isoformat():
        return None
    last = rows[-1]
    closes = [r["close"] for r in rows]
    vols = [r["vol"] for r in rows]
    prev = last["close"] - last["change"]
    s = {
        **w, "market": market,
        "close": last["close"], "change": last["change"],
        "pct": round(last["change"] / prev * 100, 2) if prev else None,
        "open": last["open"], "high": last["high"], "low": last["low"],
        "volume": last["vol"],
        "vol_ratio": round(last["vol"] / (sum(vols[-6:-1]) / 5), 2) if len(vols) >= 6 and sum(vols[-6:-1]) else None,
        "ma5": ma(closes, 5), "ma20": ma(closes, 20), "ma60": ma(closes, 60),
        "high60": max(closes[-60:]), "low60": min(closes[-60:]),
        "spark": closes[-20:],
    }
    nets = [day.get(w["code"]) for day in inst_by_day]
    if nets and nets[0]:
        s["inst"] = {"foreign": nets[0][0], "trust": nets[0][1], "dealer": nets[0][2], "total": nets[0][3]}
        valid = [n for n in nets if n]
        s["foreign_streak"] = streak([n[0] for n in valid])
        s["trust_streak"] = streak([n[1] for n in valid])
        s["foreign5"] = sum(n[0] for n in valid[:5])
        s["trust5"] = sum(n[1] for n in valid[:5])
    s["signals"] = stock_signals(s, closes)
    return s


def stock_directory() -> dict[str, dict]:
    """全部上市櫃股票與 ETF 名單 {代號: {name, market}}，給網頁搜尋用。每天更新一次。"""
    def load():
        out = {}
        for r in fetch_json("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"):
            out[r["Code"].strip()] = {"name": r["Name"].strip(), "market": "twse"}
        for r in fetch_json("https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes"):
            code = r["SecuritiesCompanyCode"].strip()
            if len(code) == 4 or code.startswith("00"):     # 股票與 ETF，排除權證、債券
                out.setdefault(code, {"name": r["CompanyName"].strip(), "market": "tpex"})
        return out
    return cached(f"directory_{date.today():%Y%m%d}", load, keep=True)


def trading_days_of(data: dict) -> list[date]:
    """從已存的資料檔取回最近 INST_DAYS 個交易日（由新到舊）。"""
    return [parse_date(r["d"]) for r in data["market"]["series"]][::-1][:INST_DAYS]


# ───────────────────────── 主流程 ─────────────────────────

def collect(target: date) -> dict:
    missing: list[str] = []
    out: dict = {"date": target.isoformat(), "generated_at": datetime.now().isoformat(timespec="seconds")}

    # 大盤
    series = taiex_series(target)
    if not series or series[-1]["d"] != target.isoformat():
        missing.append("大盤收盤")
        out["missing"] = missing
        return out
    today = series[-1]
    prev_close = today["close"] - today["change"]
    out["market"] = {
        "close": today["close"], "change": today["change"],
        "pct": round(today["change"] / prev_close * 100, 2),
        "turnover": today["turnover"],
        "turnover_avg5": round(sum(r["turnover"] for r in series[-6:-1]) / 5, 2) if len(series) >= 6 else None,
        "ma20": ma([r["close"] for r in series], 20),
        "ma60": ma([r["close"] for r in series], 60),
        "series": series[-60:],
    }
    days = [parse_date(r["d"]) for r in series][::-1][:INST_DAYS]   # 由新到舊

    # 三大法人（大盤）
    inst_hist = []
    for d in days:
        v = bfi82u(d)
        if v:
            inst_hist.append({"d": d.isoformat(), **v})
        elif d == target:
            missing.append("三大法人買賣超")
    if inst_hist and inst_hist[0]["d"] == target.isoformat():
        out["institutions"] = {
            "today": inst_hist[0],
            "history": inst_hist[::-1],
            "foreign_streak": streak([h["foreign"] for h in inst_hist]),
            "trust_streak": streak([h["trust"] for h in inst_hist]),
        }

    # 期貨
    try:
        fut = futures_tx(target)
        if fut:
            fut["basis"] = round(fut["close"] - today["close"], 2)
            out["futures"] = fut
        else:
            missing.append("台指期行情")
    except Exception as e:
        log.exception("台指期行情抓取失敗")
        missing.append(f"台指期行情（{e}）")

    try:
        hist = futures_institutions(target)
        if hist and hist[-1]["d"] == target.isoformat():
            fi = hist[-1]
            prev = hist[-2] if len(hist) > 1 else None
            out["fut_inst"] = {
                "today": fi,
                "foreign_change": fi["foreign"] - prev["foreign"] if prev else None,
                "history": hist[-20:],
            }
        else:
            missing.append("期貨三大法人未平倉")
    except Exception as e:
        log.exception("期貨法人抓取失敗")
        missing.append(f"期貨三大法人未平倉（{e}）")

    try:
        pcs = pc_ratio(target)
        ref = (out.get("futures") or {}).get("close") or today["close"]
        opt = options_txo(target, ref)
        if opt and pcs and pcs[-1]["d"] == target.isoformat():
            opt["pc_oi"] = pcs[-1]["oi"]
            opt["pc_vol"] = pcs[-1]["vol"]
            opt["pc_history"] = pcs[-20:]
            out["options"] = opt
        else:
            missing.append("選擇權未平倉")
    except Exception as e:
        log.exception("選擇權抓取失敗")
        missing.append(f"選擇權未平倉（{e}）")

    # 個股
    watch = read_watchlist()
    markets = {w["code"]: detect_market(w["code"], target) for w in watch}
    inst_by_day = inst_days(days, need_tpex="tpex" in markets.values())
    if not inst_by_day or not inst_by_day[0]:
        missing.append("個股法人買賣超")

    stocks = []
    for w in watch:
        s = stock_row(w, target, days, inst_by_day, markets[w["code"]])
        if s is None:
            missing.append(f"{w['code']} 收盤價")
        else:
            stocks.append(s)
    out["stocks"] = stocks

    # 三大法人買賣超排行（全市場）
    try:
        rank = inst_ranking(target)
        if rank:
            out["inst_rank"] = rank
        else:
            missing.append("法人買賣超排行")
    except Exception as e:
        log.exception("法人買賣超排行計算失敗")
        missing.append(f"法人買賣超排行（{e}）")

    # 美股與匯率（跑馬燈用）。Yahoo 偶爾不穩，抓不到不影響其他資料
    us = us_markets(target)
    out["us"] = us
    if read_us_watchlist() and not us:
        missing.append("美股與匯率")

    # 美日台融資。公布頻率不同（台每日晚上、日每週、美每月），一律顯示「最近一次已公布」的數字，
    # 不列入缺漏，避免下午為了等台股融資一直重試
    margin = {}
    for key, fn in (("tw", lambda: tw_margin(days)), ("us", lambda: us_margin(target)), ("jp", lambda: jp_margin(target))):
        try:
            v = fn()
            if v:
                margin[key] = v
        except Exception as e:
            log.warning("融資資料（%s）抓取失敗：%s", key, e)
    out["margin"] = margin
    out["missing"] = missing
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="台股盤後資料抓取")
    ap.add_argument("--date", help="指定日期 YYYY-MM-DD，預設今天")
    ap.add_argument("--wait", action="store_true", help="資料不齊時每 15 分鐘重試")
    ap.add_argument("--until", default="18:30", help="--wait 最晚重試到幾點（預設 18:30）")
    args = ap.parse_args()
    setup_logging("fetch")

    target = parse_date(args.date) if args.date else date.today()
    if not is_trading_day(target):
        log.info("%s 不是交易日，略過。", target)
        return 0

    deadline = datetime.combine(date.today(), datetime.strptime(args.until, "%H:%M").time())
    while True:
        log.info("開始抓取 %s 盤後資料", target)
        result = collect(target)
        missing = result.get("missing", [])
        if "market" in result:
            path = DATA_DIR / f"{target}.json"
            path.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
            log.info("已存檔：%s（%d 檔個股）", path.name, len(result.get("stocks", [])))
        if not missing:
            log.info("資料完整。")
            return 0
        log.warning("尚未取得：%s", "、".join(missing))
        if not args.wait or datetime.now() + timedelta(minutes=15) > deadline:
            return 0 if "market" in result else 2
        log.info("15 分鐘後重試……")
        time.sleep(15 * 60)


if __name__ == "__main__":
    sys.exit(main())
