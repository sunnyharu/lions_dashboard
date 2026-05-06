"""
오늘 날짜로 last_run.txt가 이미 커밋되어 있으면 skip=true 출력
→ GitHub Actions에서 이후 crawl job을 건너뜀
"""
import os
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
TODAY = datetime.now(KST).strftime("%Y-%m-%d")


def set_output(key: str, value: str):
    github_output = os.environ.get("GITHUB_OUTPUT", "")
    if github_output:
        with open(github_output, "a") as f:
            f.write(f"{key}={value}\n")
    else:
        print(f"::set-output name={key}::{value}")


def main():
    last_run_file = "last_run.txt"

    if os.path.exists(last_run_file):
        with open(last_run_file) as f:
            last_run = f.read().strip()
        if last_run == TODAY:
            print(f"오늘({TODAY}) 이미 실행됨 → 스킵")
            set_output("skip", "true")
            return

    print(f"오늘({TODAY}) 미실행 → 실행")
    set_output("skip", "false")


if __name__ == "__main__":
    main()
