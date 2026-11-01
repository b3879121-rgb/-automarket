"""把每天新產生的資料與儀表板推送到 GitHub，讓網站（GitHub Pages）跟著更新。

只會提交 data/、reports/ 與自選股清單；程式碼的修改不會被自動推送。
沒有變更時什麼都不做。結果記錄在 logs/publish_YYYY-MM.log。
"""
import subprocess
import sys
from datetime import datetime

from common import ROOT, log, setup_logging

PATHS = ["data", "reports", "watchlist.csv", "us_watchlist.csv"]


def git(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=180)


def main() -> int:
    setup_logging("publish")
    if not (ROOT / ".git").exists():
        log.info("這個資料夾不是 git repo，略過推送。")
        return 0
    git("add", "--", *PATHS)
    if git("diff", "--cached", "--quiet").returncode == 0:
        log.info("沒有新資料，不需要推送。")
        return 0
    msg = f"每日資料更新 {datetime.now():%Y-%m-%d %H:%M}"
    r = git("commit", "-m", msg)
    if r.returncode != 0:
        log.error("commit 失敗：%s", (r.stderr or r.stdout).strip())
        return 1
    r = git("push")
    if r.returncode != 0:
        log.error("推送失敗（下次執行會再試）：%s", (r.stderr or r.stdout).strip())
        return 1
    log.info("已推送到 GitHub：%s", msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
