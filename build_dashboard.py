"""把 data/*.json 套進 template.html，產生 reports/日期.html 與 reports/latest.html。

用法：
    python build_dashboard.py            產生最新一天
    python build_dashboard.py --all      所有日期全部重新產生（改過模板後用）
"""
import argparse
import json
import sys

from common import DATA_DIR, REPORT_DIR, ROOT, log, setup_logging


def render(template: str, data: dict, dates: list[str]) -> str:
    # JSON 放進 <script> 時，"</" 要跳脫，避免提前結束 script 標籤
    dump = lambda v: json.dumps(v, ensure_ascii=False).replace("</", "<\\/")
    return template.replace("__DATA__", dump(data)).replace("__DATES__", dump(dates))


def data_files() -> list:
    return sorted(DATA_DIR.glob("????-??-??.json"))


def available_dates() -> list[str]:
    """有大盤資料的日期，新到舊，給日期選單用。"""
    out = []
    for f in data_files()[::-1]:
        out.append(f.stem)
    return out


def render_date(day: str) -> str:
    template = (ROOT / "template.html").read_text(encoding="utf-8")
    data = json.loads((DATA_DIR / f"{day}.json").read_text(encoding="utf-8"))
    return render(template, data, available_dates())


def main() -> int:
    ap = argparse.ArgumentParser(description="產生台股盤後儀表板")
    ap.add_argument("--all", action="store_true", help="重新產生所有日期")
    args = ap.parse_args()
    setup_logging("build")
    return build(args.all)


def build(all_dates: bool = False) -> int:
    files = data_files()
    if not files:
        log.error("data 資料夾裡沒有資料，請先執行 fetch_after_close.py")
        return 1
    template = (ROOT / "template.html").read_text(encoding="utf-8")
    dates = available_dates()
    targets = files if all_dates else files[-1:]
    for f in targets:
        data = json.loads(f.read_text(encoding="utf-8"))
        if "market" not in data:
            log.warning("%s 沒有大盤資料，略過", f.name)
            continue
        html = render(template, data, dates)
        (REPORT_DIR / f"{f.stem}.html").write_text(html, encoding="utf-8")
        log.info("已產生 reports/%s.html", f.stem)
    latest = REPORT_DIR / f"{files[-1].stem}.html"
    if latest.exists():
        (REPORT_DIR / "latest.html").write_text(latest.read_text(encoding="utf-8"), encoding="utf-8")
        log.info("已更新 reports/latest.html")
    return 0


if __name__ == "__main__":
    sys.exit(main())
