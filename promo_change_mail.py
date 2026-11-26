# promo_change_mail.py
# 2026-11-26  Jonghyun Park w/ Claude
# updated: 2026-11-26 - v1.7: 맨 위 요약 맨 아래에 "(제외 규칙으로 뺀 링크) gr: christmas theater 1건" 회색 한 줄 추가
#   (checker 의 EXCLUDED_LINK_LOG_NAME 기록 중 기간 안 것, sitecode · 제외 규칙별 링크 수. 없으면 줄 없음. SHOW_EXCLUDED_LINKS)
# updated: 2026-11-23 — v1.6: "감지 안 됨" 을 두 줄로 나눴다 — 일정 엑셀 SCHEDULE_API_EXCLUDE_HEADER 열이 'O' 인 site 는
#   별도 줄(API 제외, 회색)로, 나머지를 먼저 굵게. checker 에서 '/offer/ 없음' 사유로 주석 처리된 site 는 (offer 없음) 으로 표시
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
  python promo_change_mail.py --test       본인에게만 [TEST] 로 보낸다 (보낸 기록 안 남김)

메일 본문 = 종류별 절(KINDS 순서: 신규 → 디자인변경) 안에 sitecode 별 비교표 (페이지 · 구분 · 링크 · 전 · 후) — 전 / 후 칸에 요소 설명(날짜 · 카드 종류 · 크기 ·
  이미지 파일명 / 텍스트)과 캡처 이미지를 나란히. 같은 캡처 쌍을 공유하는 요소는 한 행으로 묶는다.
  - 디자인변경: 같은 링크의 **앞선 날짜** 기록이 있을 때만. 같은 날 같은 링크로 함께 잡힌 다른 요소(이미지 카드 ·
    텍스트 카드 등)는 checker 가 디자인변경으로 적어도 신규로 본다.
  - 신규: 전은 "— (신규)".

  작업 스케줄러에서 checker 작업의 **두 번째 동작**으로 등록한다 — 동작은 순서대로 돌므로 감지가 끝나면 바로 메일이 나간다.
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
# 기간 안에 처음 hit 된 sitecode 는 뺀다 (checker 의 첫 hit 메일과 겹치지 않게 — 그 site 요소는 전부 신규라 중복이 된다)
SKIP_FIRST_HIT_SITES: bool = True

# ─── 기간 진행중인데 감지 안 된 site (맨 위 요약에 한 줄) ──────────────
# 캠페인 일정 엑셀 — 올해 기간 안인데 checker 최신 결과에 hit(O)가 하나도 없는 site 를 요약에 적는다. 비우면 안 적는다.
SCHEDULE_XLSX: str = ""
SCHEDULE_SHEET: str = "MASTER"
# 머리글 행 번호 (SCHEDULE_SITE_HEADER 가 있는 행)
SCHEDULE_HEADER_ROW: int = 4
SCHEDULE_SITE_HEADER: str = "Site Code"
# 올해 기간 열 (시작, 끝) — 하나라도 오늘이 들어가면 진행중. 날짜가 아닌 칸('X' 등)은 무시
SCHEDULE_PERIOD_COLS: list[tuple[str, str]] = [("H", "I"), ("J", "K"), ("L", "M")]
# API 수기 제외 표시 열의 머리글 — 이 열이 'O' 인 site 는 '감지 안 됨' 에서 별도 줄(API 제외)로 뺀다.
# 비우거나 엑셀에 그 열이 없으면 전부 한 줄로 적는다.
SCHEDULE_API_EXCLUDE_HEADER: str = "api수기제외"

# ─── 제외 규칙으로 뺀 링크 (맨 위 요약 맨 아래 회색 한 줄) ──────────────
# checker 의 EXCLUDE_LINK_KEYWORDS 로 뺀 링크(캠페인 키워드에도 걸린 것)가 기간 안에 있으면 sitecode 별로 적는다.
# False 면 안 적는다. 기간 안에 제외된 링크가 없으면 줄 자체가 없다.
SHOW_EXCLUDED_LINKS: bool = True

# ─── 메일 모양 ─────────────────────────────────────────────────────
# 제목 = checker 의 MAIL_SUBJECT_PREFIX + 아래 꼬리표. checker 접두를 그대로 쓰므로
# checker 가 실행될 때 받은 메일을 MAIL_MOVE_TO_FOLDER 로 같이 옮긴다.
SUBJECT_TAG: str = "[변경]"
# 보낸 뒤 받은편함에 도착할 때까지 최대 몇 초 기다렸다가 checker 의 MAIL_MOVE_TO_FOLDER 로 옮길지
MOVE_WAIT_SEC: int = 300
# 원본 캡처를 첨부로도 붙일지 (False 면 본문 이미지만 — 메일 용량이 절반 정도)
ATTACH_ORIGINALS: bool = True
# 요소 텍스트를 목록에 몇 자까지 적을지
TEXT_MAX_CHARS: int = 60
# 비교표에 전 / 후 캡처를 나란히 놓을 때 이미지 한 장의 가로(px)
COMPARE_IMAGE_WIDTH: int = 360
# 요소 class 에 들어 있는 문자열 → 표에 적을 카드 종류 (앞에서부터 먼저 맞는 것)
CARD_KINDS: dict[str, str] = {"image-link": "이미지 카드", "text-link": "텍스트 카드"}

# ─── 파일 ─────────────────────────────────────────────────────────
# 같은 폴더에서 찾을 checker 파일 (여러 개면 버전이 가장 높은 것)
CHECKER_GLOB: str = "promo_link_checker_v*.py"
SENT_LOG_NAME: str = "_change_mail_sent.json"
PREVIEW_HTML_NAME: str = "_change_mail_preview.html"
# pythonw(작업 스케줄러)로 돌 때 출력 · 에러를 남길 로그 (매 실행 덮어씀)
RUN_LOG_NAME: str = "_change_mail_latest.log"

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
    """같은 링크의 직전 디자인 기록 — 이번보다 **앞선 날짜**에 처음 기록된 것 중 가장 늦은 것.
    같은 날 기록은 같은 실행에서 함께 발견된 다른 요소(같은 링크의 이미지 카드 · 텍스트 카드 등)라 전으로 보지 않는다."""
    older = [v for f, v in fps.items() if f != fp and v.get("first_seen", "") < first_seen]
    return max(older, key=lambda v: v.get("first_seen", "")) if older else None


def collect(state: dict, days: set[str], sent: dict, resend: bool) -> dict[str, list[dict]]:
    """{sitecode: [요소 행, ...]} — days 안에 처음 기록됐고 아직 안 보낸 KINDS 요소. 행에 직전 기록(before)을 붙인다."""
    by_site: dict[str, list[dict]] = defaultdict(list)
    # sitecode 별 가장 이른 기록일 — 기간 안에 처음 기록된 site 는 checker 의 첫 hit 메일이 이미 알렸다
    site_first: dict[str, str] = {}
    for state_key, bucket in state.items():
        sc = state_key.partition("|")[0]
        for fps in bucket.values():
            for v in fps.values():
                d = v.get("first_seen", "")
                if d and (sc not in site_first or d < site_first[sc]):
                    site_first[sc] = d
    for state_key, bucket in state.items():
        sc, _, page = state_key.partition("|")
        if SKIP_FIRST_HIT_SITES and site_first.get(sc, "") >= min(days):
            continue
        for el_key, fps in bucket.items():
            for fp, v in fps.items():
                if v.get("kind") not in KINDS or v.get("first_seen") not in days:
                    continue
                eid = _el_id(state_key, el_key, fp)
                if eid in sent and not resend:
                    continue
                prev = _previous(fps, fp, v["first_seen"]) if v["kind"] != "new" else None
                # 앞선 날짜 기록이 없는 디자인변경 = 같은 실행의 같은 링크 다른 요소 → 신규로 본다
                kind = "new" if v["kind"] == "changed" and not prev else v["kind"]
                if kind not in KINDS:
                    continue
                by_site[sc].append({"id": eid, "page": page, "link": el_key, "kind": kind,
                                    "design": v.get("design") or {}, "capture": v.get("capture", ""),
                                    "first_seen": v["first_seen"], "before": prev})
    return by_site


def _commented_reasons(m) -> dict[str, str]:
    """checker SITECODES 에서 '#' 으로 막아 둔 sitecode → 괄호 표기. '/offer/ 없음' 사유면 'offer 없음'."""
    out = {}
    for line in (getattr(m, "SITECODES", "") or "").splitlines():
        s = line.strip()
        if not s.startswith("#"):
            continue
        parts = s.lstrip("#").split(None, 1)
        if not parts or not re.fullmatch(r"[a-z]{2}(?:_[a-z]{2,3})?|[a-z_]+", parts[0]):
            continue
        reason = parts[1] if len(parts) > 1 else ""
        if "offer" in reason.lower() and "없음" in reason:
            out[parts[0].lower()] = "offer 없음"
    return out


def undetected_sites(m, today: str) -> tuple[list[str], list[str]] | None:
    """일정 엑셀 기준 오늘 진행중인데 checker 오늘 결과에 hit(O)가 없는 sitecode — (일반, API 제외) 두 목록.
    항목 형태: "ua", "mn(offer 없음)". API 제외 = 일정 엑셀 SCHEDULE_API_EXCLUDE_HEADER 열이 'O' 인 site.
    괄호: offer 없음 = checker 에서 '/offer/ 없음' 사유로 주석 처리, 점검 목록 밖 = checker SITECODES 에 없음(그 밖의 주석 포함),
    오늘 결과 없음 = 목록엔 있는데 오늘 결과 행이 없음, 접속 실패 = 결과가 O/X 가 아닌 것(접근실패 등)만 있음.
    SCHEDULE_XLSX 가 비었거나 읽을 수 없으면 None (요약 줄을 안 적는다)."""
    if not SCHEDULE_XLSX.strip():
        return None
    try:
        import openpyxl
        from openpyxl.utils import column_index_from_string
        wb = openpyxl.load_workbook(SCHEDULE_XLSX, read_only=True, data_only=True)
        rows = list(wb[SCHEDULE_SHEET].iter_rows(min_row=SCHEDULE_HEADER_ROW, values_only=True))
        wb.close()
    except Exception as e:
        print(f"⚠ 일정 엑셀을 못 읽어 '감지 안 됨' 요약을 뺀다: {e}")
        return None
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    site_idx = header.index(SCHEDULE_SITE_HEADER)
    api_idx = header.index(SCHEDULE_API_EXCLUDE_HEADER) if SCHEDULE_API_EXCLUDE_HEADER in header else None
    if SCHEDULE_API_EXCLUDE_HEADER and api_idx is None:
        print(f"⚠ 일정 엑셀에 '{SCHEDULE_API_EXCLUDE_HEADER}' 열이 없어 API 제외 줄 없이 한 줄로 적는다.")
    pairs = [(column_index_from_string(a) - 1, column_index_from_string(b) - 1) for a, b in SCHEDULE_PERIOD_COLS]
    in_period: list[str] = []
    api_excluded: set[str] = set()
    for row in rows[1:]:
        sc = str(row[site_idx] or "").strip().lower() if site_idx < len(row) else ""
        if not sc:
            continue
        if api_idx is not None and api_idx < len(row) and str(row[api_idx] or "").strip().upper() == "O":
            api_excluded.add(sc)
        for a, b in pairs:
            s, e = (row[a] if a < len(row) else None), (row[b] if b < len(row) else None)
            if isinstance(s, datetime) and isinstance(e, datetime) and s.date().isoformat() <= today <= e.date().isoformat():
                in_period.append(sc)
                break
    # checker 오늘 결과 — 일일 리포트의 오늘 날짜 블록 첫 열(결과). 결과 CSV 는 완주하면 지워지므로 리포트를 본다.
    results: dict[str, set[str]] = defaultdict(set)
    try:
        wb = openpyxl.load_workbook(m.OUTPUT_DIR / m.DAILY_REPORT_NAME, read_only=True, data_only=True)
        rep = list(wb.active.iter_rows(values_only=True))
        wb.close()
        mmdd = today[5:7] + today[8:10]          # checker 의 run_date 형식 (%m%d)
        col = next((i for i, v in enumerate(rep[0]) if i >= m.REPORT_FIXED_COLS and str(v or "") == mmdd), None)
        if col is not None:
            for row in rep[2:]:
                if row[0] and col < len(row) and row[col] is not None:
                    results[str(row[0]).strip().lower()].add(str(row[col]).strip())
    except Exception as e:
        print(f"⚠ 일일 리포트를 못 읽음: {e}")
    if not results:
        print("⚠ checker 오늘 결과가 없어 '감지 안 됨' 요약을 뺀다.")
        return None
    checked = set(m.parse_sitecodes())
    reasons = _commented_reasons(m)
    main, api = [], []
    for sc in sorted(set(in_period)):
        got = results.get(sc)
        if sc not in checked:
            item = f"{sc}({reasons.get(sc, '점검 목록 밖')})"
        elif got is None:
            item = f"{sc}(오늘 결과 없음)"
        elif "O" not in got:
            item = sc + ("(접속 실패)" if any(x not in ("O", "X") for x in got) else "")
        else:
            continue
        (api if sc in api_excluded else main).append(item)
    return main, api


def excluded_links(m, days: list[str]) -> list[str]:
    """기간 안에 checker 가 EXCLUDE_LINK_KEYWORDS 로 뺀 링크 → ["gr: christmas theater 1건", ...].
    같은 링크는 하루 여러 번 · 여러 날 기록돼도 1건. 기록 파일이 없거나 SHOW_EXCLUDED_LINKS 가 False 면 빈 목록."""
    name = getattr(m, "EXCLUDED_LINK_LOG_NAME", "")
    if not SHOW_EXCLUDED_LINKS or not name:
        return []
    path = m.OUTPUT_DIR / name
    if not path.is_file():
        return []
    import csv
    links: dict[tuple[str, str], set[str]] = defaultdict(set)
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                if (r.get("checked_at") or "")[:10] in days:
                    links[(r["sitecode"], r["exclude_rule"])].add(r["link"])
    except Exception as e:
        print(f"⚠ 제외 링크 기록을 못 읽음: {e}")
        return []
    return [f"{sc}: {rule} {len(v)}건" for (sc, rule), v in sorted(links.items())]


def _desc(design: dict, first_seen: str, esc) -> str:
    """요소 한 줄 설명 — 날짜 · 카드 종류 · 크기 + 이미지 파일명 또는 텍스트."""
    classes = " ".join(design.get("classes") or [])
    kind = next((name for key, name in CARD_KINDS.items() if key in classes), classes or "요소")
    size = design.get("size") or []
    size_s = f" {size[0]}×{size[1]}" if len(size) == 2 else ""
    if size_s and not any(size):
        size_s = " (숨김 0×0)"
    imgs = [Path(str(u).split("?", 1)[0]).name for u in design.get("images") or []]
    text = (design.get("text") or "")[:TEXT_MAX_CHARS]
    detail = (f'<br>이미지: {esc(", ".join(imgs))}' if imgs else "") + (f'<br>텍스트: "{esc(text)}"' if text else "")
    return f'<b>{esc(first_seen)}</b> · {esc(kind)}{esc(size_s)}<span style="color:#555">{detail}</span>'


def build_html(m, by_site: dict[str, list[dict]], days: list[str],
               undetected: tuple[list[str], list[str]] | None = None,
               excluded: list[str] | None = None) -> tuple[str, list[tuple[str, Path]], dict]:
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
    # 맨 위 요약 — 종류별로 어느 site 인지 (site 옆 숫자 = 요소 수)
    lines = []
    for kind in KINDS:
        per = {sc: sum(1 for r in rows if r["kind"] == kind) for sc, rows in by_site.items()}
        per = {sc: n for sc, n in per.items() if n}
        if per:
            sites_s = ", ".join(f"{sc}({n})" for sc, n in sorted(per.items()))
            lines.append(f'<li><b style="color:{_KIND_COLORS.get(kind, "#666")}">{esc(m.KIND_LABELS.get(kind, kind))}</b> '
                         f"{len(per)}개 site: {esc(sites_s)}</li>")
    if undetected is not None:
        main_u, api_u = undetected
        # API 제외가 아닌데 감지 안 된 site 가 더 중요 → 먼저, 굵게. API 제외 site 는 다음 줄에 회색으로
        lines.append(f'<li><b style="color:#c00000">기간 진행중 · 감지 안 됨</b> {len(main_u)}개 site: '
                     f'<b>{esc(", ".join(main_u)) if main_u else "없음"}</b></li>')
        if api_u:
            lines.append(f'<li style="color:#888">(API 제외 site) 기간 진행중 · 감지 안 됨 {len(api_u)}개 site: '
                         f'{esc(", ".join(api_u))}</li>')
    if excluded:
        # 의도적으로 뺀 것이라 강조 없이 회색 (팀원이 누락으로 오해하지 않게만)
        lines.append(f'<li style="color:#888">(제외 규칙으로 뺀 링크) {esc(", ".join(excluded))}</li>')
    parts.append(f'<ul style="margin:4px 0 10px">{"".join(lines)}</ul>')
    images: list[tuple[str, Path]] = []
    cid_of: dict[Path, str] = {}       # 같은 캡처는 메일에 한 번만 싣는다
    td = 'style="border:1px solid #ccc;padding:4px 6px;vertical-align:top"'
    th = 'style="border:1px solid #ccc;padding:4px 6px;background:#f2f2f2;text-align:left"'

    shown: set[tuple[str, Path]] = set()   # (sitecode 절, 캡처) — 같은 표 안에서 같은 캡처를 두 번 그리지 않는다

    def img_html(ref: str, scope: str = "") -> str:
        if not ref:
            return ""
        path = m._capture_path(ref)
        label = esc(m._capture_label(path))
        if not path.is_file():
            return f'<span style="color:#888">캡처 파일 없음: {label}</span>'
        if (scope, path) in shown:
            return f'<span style="color:#888">↑ 위 행과 같은 캡처 ({label})</span>'
        shown.add((scope, path))
        if path not in cid_of:
            cid_of[path] = f"pcm{len(images)}"
            images.append((cid_of[path], path))
        return (f'<img src="cid:{cid_of[path]}" width="{COMPARE_IMAGE_WIDTH}"><br>'
                f'<span style="color:#888;font-size:8.5pt">{label}</span>')

    def side_cell(descs: list[str], ref: str, empty: str, scope: str = "") -> str:
        body = _ITEM_SEP.join(descs) if descs else f'<span style="color:#888">{esc(empty)}</span>'
        img = img_html(ref, scope)
        return f"<td {td}>{body}" + (f'<div style="margin-top:6px">{img}</div>' if img else "") + "</td>"

    # 종류(KINDS 순서 = 신규 → 디자인변경) 별로 큰 절을 나누고, 그 안에서 sitecode 별 표
    for kind in KINDS:
        sites = {sc: [r for r in rows if r["kind"] == kind] for sc, rows in by_site.items()}
        sites = {sc: rows for sc, rows in sites.items() if rows}
        if not sites:
            continue
        color = _KIND_COLORS.get(kind, "#666")
        parts.append(f'<h1 style="margin:24px 0 4px;font-size:15pt;color:{color}">'
                     f'{esc(m.KIND_LABELS.get(kind, kind))} {counts[kind]}건 '
                     f'<span style="font-size:10.5pt;color:#888">({len(sites)}개 site)</span></h1>')
        _site_tables(parts, sites, color, m, esc, td, th, img_html, side_cell)
    parts.append(f'<p style="margin-top:18px;color:#888">캡처 폴더: {esc(m._share_path(m.CAPTURE_DIR))}</p></div>')
    return "".join(parts), images, counts


def _site_tables(parts, by_site, color, m, esc, td, th, img_html, side_cell) -> None:
    for sc in sorted(by_site):
        # (페이지, 후 캡처, 전 캡처) 가 같은 요소는 한 행으로 묶는다
        groups: dict[tuple[str, str, str], list[dict]] = defaultdict(list)
        for r in sorted(by_site[sc], key=lambda r: (r["page"], r["link"])):
            prev = r["before"] or {}
            groups[(r["page"], r["capture"], prev.get("capture", ""))].append(r)
        parts.append(f'<h2 style="margin:16px 0 6px;font-size:12.5pt;border-bottom:2px solid {color}">{esc(sc)}</h2>'
                     '<table style="border-collapse:collapse;font-size:9.5pt;margin:4px 0 8px">'
                     f"<tr><th {th}>페이지</th><th {th}>구분 · 링크</th><th {th}>전</th><th {th}>후</th></tr>")
        keys = list(groups)
        for page, after_ref, before_ref in keys[:m.MAIL_MAX_IMAGES_PER_SITE]:
            rows = groups[(page, after_ref, before_ref)]
            links, befores, afters = [], [], []
            for r in rows:
                label = m.KIND_LABELS.get(r["kind"], r["kind"])
                link = r["link"] if r["link"].startswith("http") else "https://" + r["link"]
                links.append(f'<span style="color:{_KIND_COLORS.get(r["kind"], "#666")}">{esc(label)}</span><br>'
                             f'<a href="{esc(link, quote=True)}">{esc(r["link"])}</a>')
                afters.append(_desc(r["design"], r["first_seen"], esc))
                if r["before"]:
                    befores.append(_desc(r["before"].get("design") or {}, r["before"].get("first_seen", ""), esc))
            kind_new = all(r["kind"] == "new" for r in rows)
            parts.append(f"<tr><td {td}>{esc(page)}</td>"
                         f"<td {td}>{_ITEM_SEP.join(links)}</td>"
                         + side_cell(befores, before_ref, "— (신규)" if kind_new else "전 기록 없음", f"{color}|{sc}")
                         + side_cell(afters, after_ref, "", f"{color}|{sc}") + "</tr>")
        parts.append("</table>")
        if len(keys) > m.MAIL_MAX_IMAGES_PER_SITE:
            rest = keys[m.MAIL_MAX_IMAGES_PER_SITE:]
            parts.append(f'<p style="color:#666">비교 {len(rest)}건 더 있음 (후 캡처):<br>'
                         + "<br>".join(esc(m._capture_label(m._capture_path(a))) if a else esc(p)
                                       for p, a, _ in rest) + "</p>")


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
                    help="본인에게만 [TEST] 제목으로 보내고 보낸 기록은 남기지 않는다 (--resend 포함)")
    args = ap.parse_args()
    if args.test:
        global CHANGE_MAIL_TO
        CHANGE_MAIL_TO = ""
        args.resend = True

    m, checker_path = load_checker()
    if sys.stdout is None:           # pythonw — 콘솔이 없으니 로그 파일로
        m.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        sys.stdout = sys.stderr = open(m.OUTPUT_DIR / RUN_LOG_NAME, "w", encoding="utf-8", buffering=1)
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "promo_change_mail 시작")
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

    undetected = undetected_sites(m, days[-1])
    if undetected is not None:
        print(f"기간 진행중 · 감지 안 됨 {len(undetected[0])}개: {', '.join(undetected[0]) or '없음'}")
        if undetected[1]:
            print(f"(API 제외 site) 기간 진행중 · 감지 안 됨 {len(undetected[1])}개: {', '.join(undetected[1])}")
    excluded = excluded_links(m, days)
    if excluded:
        print(f"(제외 규칙으로 뺀 링크) {', '.join(excluded)}")
    body, images, counts = build_html(m, by_site, days, undetected, excluded)
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
        # [TEST] 는 checker 접두 뒤에 붙인다 — 제목이 접두로 시작해야 폴더 이동 대상이 된다
        subject = subject.replace(m.MAIL_SUBJECT_PREFIX, m.MAIL_SUBJECT_PREFIX + "[TEST]", 1)
    to = send(m, subject, body, images)
    print(f"📧 발송 → {to}")
    if not args.test:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        for rows in by_site.values():
            for r in rows:
                sent[r["id"]] = now
        sent_path.write_text(json.dumps(sent, ensure_ascii=False, indent=1), encoding="utf-8")
    # 이미지가 많아 받은편함 도착이 늦으므로 checker 보다 길게 기다린다
    moved = m.move_alert_mails(max(MOVE_WAIT_SEC, m.MAIL_MOVE_WAIT_SEC))
    print(f"📁 {moved}통을 '{m.MAIL_MOVE_TO_FOLDER}' 로 옮김" if moved
          else f"📁 받은편함에 아직 안 들어와 못 옮김 — 다음 checker / 이 스크립트 실행 때 같이 옮겨진다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
