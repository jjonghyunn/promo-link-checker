# promo_change_mail.py
# 2026-11-20  Jonghyun Park w/ Claude
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

메일 본문 = sitecode 별 비교표 (페이지 · 구분 · 링크 · 항목 · 전 · 후) + 후 / 전 캡처 이미지.
  - 디자인변경: 같은 요소의 직전 기록과 비교해 달라진 항목(이미지 · 텍스트 · 크기 · class)만 적고, 후 칸을 노란색으로.
  - 신규: 전은 "—", 후에 그 요소의 값.

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

# ─── 파일 ─────────────────────────────────────────────────────────
# 같은 폴더에서 찾을 checker 파일 (여러 개면 버전이 가장 높은 것)
CHECKER_GLOB: str = "promo_link_checker_v*.py"
SENT_LOG_NAME: str = "_change_mail_sent.json"
PREVIEW_HTML_NAME: str = "_change_mail_preview.html"

# ════════════════════════════════ 내부 사용 ════════════════════════════════
SCRIPT_DIR = Path(__file__).resolve().parent
_VERSION_RE = re.compile(r"_v(\d+)\.(\d+)\.py$")
_KIND_COLORS = {"new": "#c00", "changed": "#06c"}


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


def _design_fields(design: dict) -> dict[str, str]:
    """비교표에 적을 디자인 항목 — 이미지는 파일명만, 크기는 가로×세로."""
    size = design.get("size") or []
    return {
        "이미지": ", ".join(Path(str(u).split("?", 1)[0]).name for u in design.get("images") or []),
        "텍스트": (design.get("text") or "")[:TEXT_MAX_CHARS],
        "크기": f"{size[0]}×{size[1]}" if len(size) == 2 else "",
        "class": " ".join(design.get("classes") or []),
    }


def build_html(m, by_site: dict[str, list[dict]], days: list[str]) -> tuple[str, list[tuple[str, Path]], dict]:
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
    td = 'style="border:1px solid #ccc;padding:3px 6px;vertical-align:top"'
    th = 'style="border:1px solid #ccc;padding:3px 6px;background:#f2f2f2;text-align:left"'
    for sc in sorted(by_site):
        rows = sorted(by_site[sc], key=lambda r: (r["page"], KINDS.index(r["kind"]), r["link"]))
        parts.append(f'<h2 style="margin:18px 0 6px;border-bottom:2px solid #c00">{esc(sc)}</h2>'
                     '<table style="border-collapse:collapse;font-size:9.5pt;margin:4px 0 8px">'
                     f"<tr><th {th}>페이지</th><th {th}>구분</th><th {th}>링크</th><th {th}>항목</th>"
                     f"<th {th}>전</th><th {th}>후</th></tr>")
        shots: list[str] = []          # 후 캡처
        before_shots: list[str] = []   # 디자인변경의 전 캡처
        for r in rows:
            label = m.KIND_LABELS.get(r["kind"], r["kind"])
            link = r["link"] if r["link"].startswith("http") else "https://" + r["link"]
            after = _design_fields(r["design"])
            prev = r["before"]
            before = _design_fields(prev.get("design") or {}) if prev else {k: "" for k in after}
            # 신규는 후 값만, 디자인변경은 달라진 항목만 (다른 게 없으면 전체를 적는다)
            keys = [k for k in after if after[k] or before[k]]
            if prev:
                keys = [k for k in keys if after[k] != before[k]] or keys
            prev_when = f'<br><span style="color:#888">({esc(prev.get("first_seen", ""))})</span>' if prev else ""
            for i, k in enumerate(keys or ["-"]):
                head = ""
                if i == 0:
                    span = max(1, len(keys))
                    head = (f'<td {td} rowspan="{span}">{esc(r["page"])}</td>'
                            f'<td {td} rowspan="{span}"><span style="color:{_KIND_COLORS.get(r["kind"], "#666")}">'
                            f'{esc(label)}</span>{prev_when}</td>'
                            f'<td {td} rowspan="{span}"><a href="{esc(link, quote=True)}">{esc(r["link"])}</a></td>')
                b = before.get(k, "") if prev else "—"
                a = after.get(k, "")
                hl = ' background:#fff3c4' if prev and a != b else ""
                parts.append(f"<tr>{head}<td {td}>{esc(k)}</td>"
                             f'<td {td}>{esc(b) or "(없음)"}</td>'
                             f'<td style="border:1px solid #ccc;padding:3px 6px;vertical-align:top;{hl}">'
                             f'{esc(a) or "(없음)"}</td></tr>')
            if r["capture"] and r["capture"] not in shots:
                shots.append(r["capture"])
            if prev and prev.get("capture") and prev["capture"] not in before_shots:
                before_shots.append(prev["capture"])
        parts.append("</table>")
        for tag, refs in (("후", shots), ("전", before_shots)):
            for ref in refs[:m.MAIL_MAX_IMAGES_PER_SITE]:
                path = m._capture_path(ref)
                if not path.is_file():
                    parts.append(f'<p style="color:#888">[{tag}] 캡처 파일 없음: {esc(m._capture_label(path))}</p>')
                    continue
                cid = f"pcm{len(images)}"
                images.append((cid, path))
                parts.append(f'<p style="margin:10px 0 2px;color:#666"><b>[{tag}]</b> {esc(m._capture_label(path))}</p>'
                             f'<p style="margin:0"><img src="cid:{cid}" width="{m.MAIL_INLINE_IMAGE_WIDTH}"></p>')
            if len(refs) > m.MAIL_MAX_IMAGES_PER_SITE:
                rest = refs[m.MAIL_MAX_IMAGES_PER_SITE:]
                parts.append(f'<p style="color:#666">[{tag}] 캡처 {len(rest)}장 더 있음:<br>'
                             + "<br>".join(esc(m._capture_label(m._capture_path(x))) for x in rest) + "</p>")
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
