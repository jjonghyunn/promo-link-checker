# promo_change_mail.py
# 2026-11-20  Jonghyun Park w/ Claude
# updated: 2026-11-20 — v1.1: 비교표를 전 / 후 캡처 이미지를 한 행에 나란히 놓는 형식으로 변경 (속성값 비교 제거)
# updated: 2026-11-20 — v1.0: 신규 · 디자인변경 클릭 요소 알림 메일 (promo_link_checker 와 같은 폴더에 두고 실행)
# (이전 버전 이력은 git history / GitHub Releases 참조 — 헤더에는 최근 2개 항목만 남긴다)
r"""
promo_change_mail — promo_link_checker 가 새로 캡처한 클릭 요소(신규 · 디자인변경)를 모아 메일로 알린다.

promo_link_checker 본체의 메일은 "그 sitecode 가 처음 hit 됐을 때" 한 번만 간다. 이미 hit 기록이 있는
sitecode 에서 새 배너가 붙거나 디자인이 바뀐 것은 캡처 · 리포트에만 남으므로, 그걸 따로 알리는 도구다.

동작
  1) 같은 폴더의 promo_link_checker_v*.py (최신 버전 1개) 를 불러와 설정값(OUTPUT_DIR · CAPTURE_DIR ·
     메일 제목 접두 · 폴더 이동 등)을 그대로 쓴다 — 캠페인 사본마다 이 파일만 복사하면 된다.
  2) output/_capture_state.json 에서 SINCE_DAYS 안에 처음 기록된 요소 중 KINDS(신규 · 디자인변경)만 고른다.
  3) 이미 보낸 요소(output/_change_mail_sent.json)는 빼고, 남은 게 있으면 메일 1통으로 묶어 보낸다.
  4) 보낸 요소를 기록한다 → 하루에 몇 번 돌려도 같은 요소는 다시 안 간다.

실행
  python promo_change_mail.py              보낼 게 있으면 보낸다
  python promo_change_mail.py --dry-run    보내지 않고 대상 · 본문만 확인 (output/_change_mail_preview.html)
  python promo_change_mail.py --day 2026-11-20 --resend   그날 요소를 보낸 기록과 상관없이 다시 보낸다
  python promo_change_mail.py --test       본인에게만 [TEST] 로 보낸다 (보낸 기록 · 폴더 이동 없음)

메일 본문 = sitecode 별 비교표 (페이지 · 요소 · 전 캡처 · 후 캡처) — 전 / 후 이미지를 한 행에 나란히.
  - 같은 캡처 쌍을 공유하는 요소는 한 행으로 묶는다 (요소 칸에 구분 · 링크 · 텍스트).
  - 디자인변경: 전 = 같은 링크의 직전 기록 캡처. 신규: 전은 "— (신규)".

  promo_link_checker 실행 직후에 돌리면 된다 (작업 스케줄러에서 checker 다음 동작으로 추가하거나 별도 작업).
"""
from __future__ import annotations

import argparse
import html
import importlib.util
import json
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

# ════════════════════════════ 사용자가 바꿔야 하는 부분 ════════════════════════════

# ─── 받는 사람 ─────────────────────────────────────────────────────
# 여러 명이면 세미콜론. 비우면 Outlook 기본 계정 = 자기 자신에게 보낸다.
# (checker 의 MAIL_TO 와 따로 둔다 — 첫 hit 는 팀, 변경 알림은 본인만 같은 식으로 나눌 수 있게)
CHANGE_MAIL_TO: str = ""

# ─── 대상 ─────────────────────────────────────────────────────────
# 알릴 요소 종류 — checker 의 KIND_LABELS 키 ("new" 신규 / "changed" 디자인변경 / "recapture" 재캡처)
KINDS: tuple[str, ...] = ("new", "changed")
# 오늘 포함 며칠 안에 처음 기록된 요소를 볼지 (1 = 오늘만). 보낸 기록이 있어 겹쳐도 다시 가지 않는다.
SINCE_DAYS: int = 1

# ─── 메일 모양 ─────────────────────────────────────────────────────
# 제목 = checker 의 MAIL_SUBJECT_PREFIX + 아래 꼬리표. checker 접두를 그대로 쓰므로
# checker 가 실행될 때 받은 메일을 MAIL_MOVE_TO_FOLDER 로 같이 옮긴다.
SUBJECT_TAG: str = "[변경]"
# 원본 캡처를 첨부로도 붙일지 (False 면 본문 이미지만 — 메일 용량이 절반 정도)
ATTACH_ORIGINALS: bool = True
# 요소 텍스트를 목록에 몇 자까지 적을지
TEXT_MAX_CHARS: int = 60
# 비교표에 전 / 후 캡처를 나란히 놓을 때 이미지 한 장의 가로(px)
COMPARE_IMAGE_WIDTH: int = 360

# ─── 파일 ─────────────────────────────────────────────────────────
# 같은 폴더에서 찾을 checker 파일 (여러 개면 버전이 가장 높은 것)
CHECKER_GLOB: str = "promo_link_checker_v*.py"
SENT_LOG_NAME: str = "_change_mail_sent.json"
PREVIEW_HTML_NAME: str = "_change_mail_preview.html"

# ════════════════════════════════ 내부 사용 ════════════════════════════════
SCRIPT_DIR = Path(__file__).resolve().parent
_VERSION_RE = re.compile(r"_v(\d+)\.(\d+)\.py$")
_KIND_COLORS = {"new": "#c00", "changed": "#06c"}
_ITEM_SEP = '<hr style="border:0;border-top:1px dashed #ccc">'


def load_checker():
    """같은 폴더의 promo_link_checker 를 모듈로 불러온다 (설정값 · 헬퍼 재사용)."""
    found = [p for p in SCRIPT_DIR.glob(CHECKER_GLOB) if _VERSION_RE.search(p.name)]
    if not found:
        sys.exit(f"❌ {SCRIPT_DIR} 에 {CHECKER_GLOB} 가 없습니다.")
    path = max(found, key=lambda p: tuple(int(x) for x in _VERSION_RE.search(p.name).groups()))
    spec = importlib.util.spec_from_file_location("promo_link_checker", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod, path


def _el_id(state_key: str, el_key: str, fp: str) -> str:
    return f"{state_key}|{el_key}|{fp}"


def load_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}


def _previous(fps: dict, fp: str, first_seen: str) -> dict | None:
    """같은 요소의 직전 디자인 기록 (이번 지문을 뺀 것 중 first_seen 이 이번 이하에서 가장 늦은 것)."""
    older = [v for f, v in fps.items() if f != fp and v.get("first_seen", "") <= first_seen]
    return max(older, key=lambda v: v.get("first_seen", "")) if older else None


def collect(state: dict, days: set[str], sent: dict, resend: bool) -> dict[str, list[dict]]:
    """{sitecode: [요소 행, ...]} — days 안에 처음 기록됐고 아직 안 보낸 KINDS 요소. 행에 직전 기록(before)을 붙인다."""
    by_site: dict[str, list[dict]] = defaultdict(list)
    for state_key, bucket in state.items():
        sc, _, page = state_key.partition("|")
        for el_key, fps in bucket.items():
            for fp, v in fps.items():
                if v.get("kind") not in KINDS or v.get("first_seen") not in days:
                    continue
                eid = _el_id(state_key, el_key, fp)
                if eid in sent and not resend:
                    continue
                prev = _previous(fps, fp, v["first_seen"]) if v["kind"] != "new" else None
                by_site[sc].append({"id": eid, "page": page, "link": el_key, "kind": v["kind"],
                                    "design": v.get("design") or {}, "capture": v.get("capture", ""),
                                    "first_seen": v["first_seen"], "before": prev})
    return by_site


def build_html(m, by_site: dict[str, list[dict]], days: list[str]) -> tuple[str, list[tuple[str, Path]], dict]:
    """sitecode 별 표 — 한 행 = 같은 캡처 쌍(전 · 후)을 공유하는 요소 묶음. 전 / 후 캡처 이미지를 나란히 놓는다."""
    esc = html.escape
    counts = defaultdict(int)
    for rows in by_site.values():
        for r in rows:
            counts[r["kind"]] += 1
    summary = " · ".join(f"{m.KIND_LABELS.get(k, k)} {counts[k]}" for k in KINDS if counts[k])
    period = days[0] if len(days) == 1 else f"{days[0]} ~ {days[-1]}"
    parts = ['<div style="font-family:Segoe UI,Malgun Gothic,sans-serif;font-size:10.5pt">',
             f"<p>{esc(period)} 새로 캡처된 클릭 요소 <b>{sum(counts.values())}건</b> ({esc(summary)}), "
             f"sitecode <b>{len(by_site)}개</b></p>"]
    images: list[tuple[str, Path]] = []
    cid_of: dict[Path, str] = {}       # 같은 캡처는 메일에 한 번만 싣는다
    td = 'style="border:1px solid #ccc;padding:4px 6px;vertical-align:top"'
    th = 'style="border:1px solid #ccc;padding:4px 6px;background:#f2f2f2;text-align:left"'

    def img_cell(ref: str, empty: str) -> str:
        if not ref:
            return f'<td {td}><span style="color:#888">{esc(empty)}</span></td>'
        path = m._capture_path(ref)
        label = esc(m._capture_label(path))
        if not path.is_file():
            return f'<td {td}><span style="color:#888">캡처 파일 없음<br>{label}</span></td>'
        if path not in cid_of:
            cid_of[path] = f"pcm{len(images)}"
            images.append((cid_of[path], path))
        return (f'<td {td}><img src="cid:{cid_of[path]}" width="{COMPARE_IMAGE_WIDTH}"><br>'
                f'<span style="color:#888;font-size:8.5pt">{label}</span></td>')

    for sc in sorted(by_site):
        # (페이지, 후 캡처, 전 캡처) 가 같은 요소는 한 행으로 묶는다
        groups: dict[tuple[str, str, str], list[dict]] = defaultdict(list)
        for r in sorted(by_site[sc], key=lambda r: (r["page"], KINDS.index(r["kind"]), r["link"])):
            prev = r["before"] or {}
            groups[(r["page"], r["capture"], prev.get("capture", ""))].append(r)
        parts.append(f'<h2 style="margin:18px 0 6px;border-bottom:2px solid #c00">{esc(sc)}</h2>'
                     '<table style="border-collapse:collapse;font-size:9.5pt;margin:4px 0 8px">'
                     f"<tr><th {th}>페이지</th><th {th}>요소</th><th {th}>전</th><th {th}>후</th></tr>")
        keys = list(groups)
        for page, after_ref, before_ref in keys[:m.MAIL_MAX_IMAGES_PER_SITE]:
            items = []
            for r in groups[(page, after_ref, before_ref)]:
                label = m.KIND_LABELS.get(r["kind"], r["kind"])
                link = r["link"] if r["link"].startswith("http") else "https://" + r["link"]
                text = (r["design"].get("text") or "")[:TEXT_MAX_CHARS]
                when = f' <span style="color:#888">(전 {esc(r["before"].get("first_seen", ""))})</span>' if r["before"] else ""
                items.append(f'<span style="color:{_KIND_COLORS.get(r["kind"], "#666")}">{esc(label)}</span>{when}<br>'
                             f'<a href="{esc(link, quote=True)}">{esc(r["link"])}</a>'
                             + (f'<br><span style="color:#555">{esc(text)}</span>' if text else ""))
            kind_new = all(r["kind"] == "new" for r in groups[(page, after_ref, before_ref)])
            parts.append(f"<tr><td {td}>{esc(page)}</td>"
                         f"<td {td}>{_ITEM_SEP.join(items)}</td>"
                         + img_cell(before_ref, "— (신규)" if kind_new else "전 캡처 없음")
                         + img_cell(after_ref, "캡처 없음") + "</tr>")
        parts.append("</table>")
        if len(keys) > m.MAIL_MAX_IMAGES_PER_SITE:
            rest = keys[m.MAIL_MAX_IMAGES_PER_SITE:]
            parts.append(f'<p style="color:#666">비교 {len(rest)}건 더 있음 (후 캡처):<br>'
                         + "<br>".join(esc(m._capture_label(m._capture_path(a))) if a else esc(p)
                                       for p, a, _ in rest) + "</p>")
    parts.append(f'<p style="margin-top:18px;color:#888">캡처 폴더: {esc(m._share_path(m.CAPTURE_DIR))}</p></div>')
    return "".join(parts), images, counts


def send(m, subject: str, body: str, images: list[tuple[str, Path]]) -> str:
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    mail = outlook.CreateItem(m._OL_MAIL_ITEM)
    mail.To = CHANGE_MAIL_TO.strip() or m._outlook_self_address(ns)
    mail.Subject = subject
    mail.HTMLBody = body
    for cid, path in images:
        inline = mail.Attachments.Add(str(path))
        inline.PropertyAccessor.SetProperty(m._MAPI_ATTACH_CONTENT_ID, cid)
        if ATTACH_ORIGINALS:
            mail.Attachments.Add(str(path))
    to = mail.To
    mail.Send()
    return to


def main() -> int:
    ap = argparse.ArgumentParser(description="promo_link_checker 신규 · 디자인변경 요소 알림 메일")
    ap.add_argument("--dry-run", action="store_true", help="보내지 않고 대상 · 본문 미리보기만")
    ap.add_argument("--day", help="이 날짜(YYYY-MM-DD) 하루만 본다 (기본: 오늘부터 SINCE_DAYS 일)")
    ap.add_argument("--resend", action="store_true", help="보낸 기록과 상관없이 다시 보낸다")
    ap.add_argument("--test", action="store_true",
                    help="본인에게만 [TEST] 제목으로 보내고 보낸 기록 · 폴더 이동은 하지 않는다 (--resend 포함)")
    args = ap.parse_args()
    if args.test:
        global CHANGE_MAIL_TO
        CHANGE_MAIL_TO = ""
        args.resend = True

    m, checker_path = load_checker()
    if args.day:
        days = [args.day]
    else:
        today = datetime.now().date()
        days = [(today - timedelta(days=i)).isoformat() for i in range(max(1, SINCE_DAYS) - 1, -1, -1)]
    state = load_json(m.OUTPUT_DIR / m.CAPTURE_STATE_NAME)
    sent_path = m.OUTPUT_DIR / SENT_LOG_NAME
    sent = load_json(sent_path)
    by_site = collect(state, set(days), sent, args.resend)
    n = sum(len(v) for v in by_site.values())
    print(f"checker: {checker_path.name} · 기간 {days[0]} ~ {days[-1]} · 대상 {n}건 / sitecode {len(by_site)}개")
    if not n:
        print("보낼 요소 없음.")
        return 0

    body, images, counts = build_html(m, by_site, days)
    summary = " · ".join(f"{m.KIND_LABELS.get(k, k)} {counts[k]}" for k in KINDS if counts[k])
    subject = f"{m.MAIL_SUBJECT_PREFIX}{SUBJECT_TAG} {days[-1][5:]} {summary} ({len(by_site)}개 site: {', '.join(sorted(by_site))})"
    print(subject)
    print(f"이미지 {len(images)}장")

    if args.dry_run:
        preview = m.OUTPUT_DIR / PREVIEW_HTML_NAME
        preview.write_text(body, encoding="utf-8")
        print(f"(dry-run) 미리보기: {preview}")
        return 0

    if args.test:
        subject = "[TEST]" + subject
    to = send(m, subject, body, images)
    print(f"📧 발송 → {to}")
    if args.test:
        return 0
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    for rows in by_site.values():
        for r in rows:
            sent[r["id"]] = now
    sent_path.write_text(json.dumps(sent, ensure_ascii=False, indent=1), encoding="utf-8")
    moved = m.move_alert_mails(m.MAIL_MOVE_WAIT_SEC)
    if moved:
        print(f"📁 {moved}통을 '{m.MAIL_MOVE_TO_FOLDER}' 로 옮김")
    return 0


if __name__ == "__main__":
    sys.exit(main())
