"""
뉴스이슈 AI요약 + 카페트렌드 다이제스트 백필

Claude 호출이 실패한 기간(요약이 원문 앞 120자로 대체되고 다이제스트가 비어 있는 날)을
다시 생성한다. 2026-09-01 ~ 2026-10-01 크레딧 소진 구간 복구용으로 만들었다.

  python3 backfill_ai_summary.py --from 2026.09.01 --to 2026.10.01 --dry-run
  python3 backfill_ai_summary.py --from 2026.09.01 --to 2026.10.01
  python3 backfill_ai_summary.py --from 2026.09.01 --to 2026.10.01 --only digest

- 뉴스이슈: 기간 안의 행 AI요약을 전부 다시 생성해 제자리에 덮어쓴다 (행 추가 없음)
- 카페트렌드: 기간 안에서 다이제스트가 없는 날만 채운다. 원래는 그날 수집한 카페 글
  최대 50건으로 만들지만 시트에는 TOP10만 남아 있어, 날짜 기준 최근 3일 TOP10 글로 근사한다
- Claude 호출이 한 번이라도 실패하면 즉시 중단한다 (원문 대체값으로 덮어쓰지 않는다)
"""
import argparse
import sys
import time
from datetime import datetime, timedelta

import crawl_naver_news as c

DIGEST_NOTE = "\n\n※ 백필: 당시 카페 TOP10 글 기준으로 재생성"
NEWS_PROMPT = "다음 기사를 한 줄로 요약해줘. 상품 판매·출시 정보가 있으면 꼭 포함해.\n\n{body}"
CAFE_PROMPT = "다음 카페 글의 주요 내용을 한 줄로 요약해줘.\n제목: {title}\n본문: {body}"


def parse_args():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="start", required=True, help="YYYY.MM.DD")
    ap.add_argument("--to", dest="end", required=True, help="YYYY.MM.DD")
    ap.add_argument("--only", choices=["news", "digest"], help="한쪽만 실행")
    ap.add_argument("--sheet", default=c.SPREADSHEET_ID, help="대상 스프레드시트 ID")
    ap.add_argument("--dry-run", action="store_true", help="대상만 출력하고 쓰지 않음")
    return ap.parse_args()


def norm(d: str) -> str:
    return d.strip().replace("-", ".")


def claude_or_abort(prompt: str, max_tokens: int = 300) -> str:
    out = c.call_claude(prompt, max_tokens=max_tokens)
    if not out:
        sys.exit(f"\n중단: Claude 호출 실패 — {c.CLAUDE_LAST_ERROR or '빈 응답'}\n"
                 f"이미 쓴 행은 유지된다. 키·크레딧 확인 후 같은 명령으로 다시 실행하면 이어서 덮어쓴다.")
    return out


def backfill_news(sh, start, end, dry):
    ws = sh.worksheet(c.SHEET_NEWS)
    rows = ws.get_all_values()
    h = rows[0]
    di, si, ti, ki, li = (h.index(x) for x in ("날짜", "출처", "제목", "AI요약", "링크"))
    col = chr(ord("A") + ki)

    targets = [(i, r) for i, r in enumerate(rows[1:], start=2)
               if r and start <= norm(r[di]) <= end]
    print(f"[뉴스이슈] 대상 {len(targets)}행 ({start} ~ {end})")
    if dry:
        for i, r in targets[:5]:
            print(f"  {i}행 {norm(r[di])} [{r[si]}] {r[ti][:30]} | 현재: {r[ki][:40]}")
        return

    buf = []
    for n, (i, r) in enumerate(targets, 1):
        body = c.fetch_text(r[li]) or r[ki]          # 본문 못 가져오면 기존 원문 조각으로
        if "카페" in r[si]:
            prompt = CAFE_PROMPT.format(title=r[ti], body=body[:1500])
        else:
            prompt = NEWS_PROMPT.format(body=body[:2000])
        buf.append({"range": f"{col}{i}", "values": [[claude_or_abort(prompt)]]})
        if len(buf) >= 20 or n == len(targets):      # 중간 실패 대비 20행마다 저장
            ws.batch_update(buf, value_input_option="RAW")
            print(f"  {n}/{len(targets)} 저장")
            buf = []
        time.sleep(0.3)


def backfill_digest(sh, start, end, dry):
    ws_news = sh.worksheet(c.SHEET_NEWS)
    news = ws_news.get_all_values()
    h = news[0]
    di, si, ti, ki = (h.index(x) for x in ("날짜", "출처", "제목", "AI요약"))
    cafe = [(norm(r[di]), r[ti], r[ki]) for r in news[1:] if r and "카페" in r[si]]

    ws = sh.worksheet(c.SHEET_DIGEST)
    have = {norm(r[0]) for r in ws.get_all_values()[1:] if r}

    day = datetime.strptime(start, "%Y.%m.%d")
    last = datetime.strptime(end, "%Y.%m.%d")
    missing = []
    while day <= last:
        d = day.strftime("%Y.%m.%d")
        if d not in have:
            lo = (day - timedelta(days=3)).strftime("%Y.%m.%d")
            posts = [p for p in cafe if lo <= p[0] <= d][:50]
            missing.append((d, posts))
        day += timedelta(days=1)

    print(f"[카페트렌드] 다이제스트 없는 날 {len(missing)}일")
    if dry:
        for d, posts in missing[:5]:
            print(f"  {d}: 근거 글 {len(posts)}건")
        return

    for d, posts in missing:
        if not posts:
            print(f"  {d}: 근거 글 없음 → 건너뜀")
            continue
        text = "\n".join(f"- {t}: {s[:80]}" for _, t, s in posts)
        digest = claude_or_abort(
            "다음은 삼성 라이온즈 팬 카페(사자사랑방)의 최근 7일간 글 목록입니다.\n"
            f"팬들이 주로 관심 갖는 상품, 이슈, 반응 등을 3~5줄로 정리해줘.\n\n{text}",
            max_tokens=500,
        )
        ws.append_row([d, digest + DIGEST_NOTE])
        print(f"  {d}: 저장 (근거 {len(posts)}건)")
        time.sleep(0.3)


def main():
    a = parse_args()
    start, end = norm(a.start), norm(a.end)
    sh = c.get_gspread_client().open_by_key(a.sheet)
    print(f"대상 시트: {sh.title} ({a.sheet})" + ("  [DRY-RUN]" if a.dry_run else ""))

    # 뉴스 요약을 먼저 고쳐두면 다이제스트 근거 글도 실제 요약으로 들어간다
    if a.only in (None, "news"):
        backfill_news(sh, start, end, a.dry_run)
    if a.only in (None, "digest"):
        backfill_digest(sh, start, end, a.dry_run)
    print("\n완료" if not a.dry_run else "\nDRY-RUN 종료 (쓰기 없음)")


if __name__ == "__main__":
    main()
