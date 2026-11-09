#!/usr/bin/env python3
"""세일즈보드 시트 → 오늘 보낼 인스타그램 DM 큐 + 초안 생성.

사용법:
  python3 dm_queue.py <세일즈시트.xlsx | 세일즈보드.csv> [--today 2026-10-28] [--limit 20] [--json]

- xlsx면 '세일즈보드' 탭을 읽고, csv면 세일즈보드 탭을 내보낸 파일로 본다.
- 규칙은 시트 '기준값' 탭의 운영 규칙을 그대로 따른다.
  1차 발송 → 3영업일 무응답 시 후속 1회 → 7일 무응답 시 채널 변경 1회 → 종결
- 초안은 templates/ 폴더의 텍스트를 채운 것. 발송 전 반드시 사람이 확인한다.
"""
import argparse
import csv
import datetime as dt
import json
import re
import sys
from pathlib import Path

TEMPLATE_DIR = Path(__file__).resolve().parent.parent / "templates"
FOLLOWUP_BUSINESS_DAYS = 3
CHANNEL_SWITCH_DAYS = 7

# 큐에 넣을 단계 → 템플릿 파일
STAGE_TEMPLATES = {
    "1차 DM": "first_dm.txt",
    "후속 DM": "followup_dm.txt",
}


def load_rows(path):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        with path.open(encoding="utf-8-sig") as f:
            return list(csv.DictReader(f))
    import openpyxl

    ws = openpyxl.load_workbook(path, data_only=True, read_only=True)["세일즈보드"]
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h else "" for h in next(rows)]
    return [dict(zip(header, r)) for r in rows if r and r[0]]


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    m = re.match(r"(\d{4})[-./]\s*(\d{1,2})[-./]\s*(\d{1,2})", str(v or ""))
    return dt.date(*map(int, m.groups())) if m else None


def business_days_between(start, end):
    """start 다음날부터 end까지의 평일 수 (공휴일은 반영하지 않음)."""
    days, d = 0, start
    while d < end:
        d += dt.timedelta(days=1)
        if d.weekday() < 5:
            days += 1
    return days


def code(v):
    """'5. 1차 발송 완료' → '5', '- 해당 없음' → '-'"""
    return str(v or "").strip().split(".")[0].strip() or "-"


def handle_from(url):
    m = re.search(r"instagram\.com/([^/?#]+)", str(url or ""))
    return m.group(1) if m else ""


def honorific(name):
    name = re.sub(r"\s*\(.*?\)", "", str(name)).strip()
    return f"{name}님"


def field_label(field):
    """'마케팅 · 퍼스널브랜딩 전문가·강연' → '퍼스널브랜딩 전문가·강연'"""
    parts = [p.strip() for p in str(field or "").split(" · ") if p.strip()]
    return parts[-1] if parts else ""


def classify(row, today):
    """(단계, 사유) 반환. 단계가 None이면 큐에서 제외."""
    send, reply = code(row.get("발송·접촉 상태")), code(row.get("회신·진행 상태"))
    next_action = str(row.get("다음 조치") or "")
    last = to_date(row.get("최근 액션일"))
    due = to_date(row.get("다음 조치 기한"))
    has_handle = bool(handle_from(row.get("연락처·계정")))

    if send in {"3", "8"} or reply in {"7", "8"} or "보류" in next_action:
        return None, "중단·보류"
    if reply in {"3", "4", "5", "6"}:
        return "직접 대응", "회신 받은 건 — 자동 후속 금지, 대화 맥락 보고 사람이 답장"
    if send == "4":
        return "발송 여부 확인", "DM 스레드 열어 실제 발송 기록 확인 → 5 또는 2로 확정"
    if send == "1" or not has_handle:
        return "계정 확인", "인스타 계정 URL 없음"
    if send in {"0", "2"}:
        return "1차 DM", "미접촉/초안 상태"
    if send == "5" and last:
        bd = business_days_between(last, today)
        if bd >= FOLLOWUP_BUSINESS_DAYS or (due and due <= today):
            return "후속 DM", f"1차 발송 후 {bd}영업일 무응답"
        return None, f"후속까지 {FOLLOWUP_BUSINESS_DAYS - bd}영업일 남음"
    if send == "6" and last and (today - last).days >= CHANNEL_SWITCH_DAYS:
        return "채널 변경·종결 판단", "후속까지 7일+ 무응답"
    return None, "대기"


def render(stage, row):
    tpl_name = STAGE_TEMPLATES.get(stage)
    if not tpl_name:
        return ""
    tpl = (TEMPLATE_DIR / tpl_name).read_text(encoding="utf-8").strip()
    return tpl.format(
        호칭=honorific(row.get("대상/업체")),
        분야=field_label(row.get("전문 분야")),
        주제=str(row.get("협업·제안 주제") or "").strip(),
    )


PRIORITY_ORDER = {"A": 0, "B": 1, "C": 2}
STAGE_ORDER = ["직접 대응", "1차 DM", "후속 DM", "발송 여부 확인", "계정 확인", "채널 변경·종결 판단"]


def build_queue(rows, today, limit):
    items = []
    for row in rows:
        if "instagram" not in str(row.get("연락 채널") or "").lower():
            continue
        stage, reason = classify(row, today)
        if not stage:
            continue
        items.append({
            "id": row.get("ID"),
            "name": row.get("대상/업체"),
            "priority": row.get("우선순위") or "",
            "handle": handle_from(row.get("연락처·계정")),
            "url": row.get("연락처·계정") or "",
            "stage": stage,
            "reason": reason,
            "topic": row.get("협업·제안 주제") or "",
            "draft": render(stage, row),
        })
    items.sort(key=lambda i: (STAGE_ORDER.index(i["stage"]), PRIORITY_ORDER.get(i["priority"], 9), str(i["id"])))
    sendable = [i for i in items if i["draft"]][:limit]
    others = [i for i in items if not i["draft"]]
    return sendable, others


def print_markdown(sendable, others, today):
    print(f"# 인스타 DM 큐 — {today.isoformat()}\n")
    print(f"## 발송 대기 {len(sendable)}건 (컨펌 후 발송)\n")
    for n, i in enumerate(sendable, 1):
        print(f"### {n}. [{i['id']}] {i['name']} · @{i['handle']} · {i['stage']} · 우선순위 {i['priority']}")
        print(f"- 사유: {i['reason']}")
        print(f"- 프로필: {i['url']}\n")
        print("```text\n" + i["draft"] + "\n```\n")
    if others:
        print(f"## 사람 확인 필요 {len(others)}건\n")
        print("| ID | 대상 | 계정 | 단계 | 사유 |\n|---|---|---|---|---|")
        for i in others:
            print(f"| {i['id']} | {i['name']} | @{i['handle'] or '-'} | {i['stage']} | {i['reason']} |")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("sheet")
    ap.add_argument("--today", type=dt.date.fromisoformat, default=dt.date.today())
    ap.add_argument("--limit", type=int, default=20, help="하루 발송 상한 (계정 보호)")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    sendable, others = build_queue(load_rows(args.sheet), args.today, args.limit)
    if args.json:
        json.dump({"today": args.today.isoformat(), "send": sendable, "review": others},
                  sys.stdout, ensure_ascii=False, indent=2)
    else:
        print_markdown(sendable, others, args.today)


if __name__ == "__main__":
    main()
