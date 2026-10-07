# promo_link_checker  
<sub>2026-10-07  Jonghyun Park w/ Claude</sub>  
국가별 **home · 프로모션(offer) 페이지의 클릭 가능한 요소** 중에 타겟 캠페인 링크가 걸려 있는지 매일 확인하고,
결과를 `_daily_report.xlsx` 에 하루 1블록씩 누적하는 도구. 캡처는 **처음 보는 hit 요소, 또는 디자인이
바뀐 요소가 있을 때만 1회** 찍는다(매일 누적 캡처하지 않음).

- 단일 파일 `promo_link_checker_v1.2.py` — 설정은 전부 상단 `사용자가 바꿔야 하는 부분` 블록.

---

## 1. 준비

- Windows / Python 3.10+ / Chrome 설치
- `pip install playwright openpyxl`
- `python -m playwright install chromium` (최초 1회)

## 2. 설정 (파일 상단)

### KEYWORDS — 감지 규칙
따옴표 없이 **한 줄에 규칙 하나**.

```python
KEYWORDS = """
summer sale
"""
```

- 링크의 **URL 경로를 `/` 로 쪼갠 구간 하나 안에** 규칙 단어가 **전부** 들어 있어야 hit (AND).
  구간 안에서의 순서·사이 단어·구분자(`- _ .`)·대소문자는 상관없다.
- 줄끼리는 **OR** — 어느 규칙이든 하나만 걸리면 hit.

| 링크 | 결과 | 이유 |
|---|---|---|
| `…/ae/summer-sale/` | ✅ hit | 한 구간에 summer·sale |
| `…/ae/offer/summer-big-sale/` | ✅ hit | 사이에 낀 단어 무관 |
| `…/ae/offer/SummerSale2026/` | ✅ hit | 붙여 써도 부분일치 |
| `…/model-x/buy/` (텍스트 "Summer Sale") | ❌ | **링크 텍스트·콜값 속성은 매칭에 안 쓴다** (리포트 설명용) |
| `partner.example.com/summer/megasales/` | ❌ | 두 단어가 **다른 구간** |
| `…/offer/?cid=summer-sale` | ❌ | **쿼리 파라미터·#fragment 제외** |
| `…/assets/terms_summer_sale_2026.pdf` | ❌ | **`.pdf` 링크 제외** (`EXCLUDE_LINK_EXTENSIONS`) — 경로에 캠페인명이 든 약관·고지 PDF 오탐 방지 |
| `…/offer/pay-sale/` | ❌ | **도메인(host) 제외** — 넣으면 모든 링크가 브랜드명을 가져 `sale` 만으로 걸린다 |

### CHECK_PAGES — 볼 페이지
따옴표 없이 한 줄에 하나. 쓸 수 있는 값은 `PAGE_SEGMENTS` 의 키(`home`, `offer`). 빼려면 줄을 지우거나 `#`.

```python
CHECK_PAGES = """
home
offer
"""
```

- `home` = `/{sitecode}/`, `offer` = `/{sitecode}/offer/` (사이트별 예외는 `SITE_OVERRIDES`).
- 실행 때만 바꾸려면 `--pages home`.

### SITECODES — 대상 국가
따옴표·쉼표 없이 **한 줄에 sitecode 하나**. `#` 로 시작하는 줄은 무시된다.

2026-09-29 전수 점검 결과(홈 92개에 `/offer/` 를 붙여 실제 브라우저로 확인):

| 구분 | sitecode | 처리 |
|---|---|---|
| 정상 | 86개 (`/offer/`) + `hq` (`/hq/shop/`) | 대상 |
| 슬러그가 다름 | `ru` → `/ru/hero-products/` (`/ru/offer/` 는 홈으로 리다이렉트) | `SITE_OVERRIDES` 로 대상 |
| offer 페이지 없음 | `iran`, `mn`, `ps` (홈으로 리다이렉트, `/offers/`·`/promo/`·`/deals/` 등도 없음) | `#` 주석 |
| 운영 중단 | `ge` | 목록에서 제외 |

→ 현재 **88개국**.

### SITE_OVERRIDES — 기본 패턴을 안 따르는 사이트
| 키 | 의미 |
|---|---|
| `base_domain` | 그 사이트 전용 도메인 (`cn` → `company_name.com.cn`) |
| `include_sitecode` | `False` 면 URL 에 sitecode 를 안 넣는다 (`cn`) |
| `segments` | 페이지별로 그 사이트에서 쓸 슬러그 — `{"offer": "shop"}` (`hq`), `{"offer": "hero-products"}` (`ru`) |

### 그 외
- `EXCLUDE_GLOBAL_UI` — `True` 면 GNB/Footer 안 링크는 hit 에서 뺀다. 기본 `False`(페이지 전체)이고,
  리포트에는 `[GNB/Footer]` 태그가 붙어 구분된다.
- `EXCLUDE_HIT_SELECTOR` — 이 선택자 안(또는 자신)의 링크는 키워드가 맞아도 hit 에서 뺀다. 기본값은 채팅봇 위젯
  프로모 배너(`[class*='rcw-promo']`) — 닫힌 위젯 안이라 화면에 안 보이는데, 배너 이미지만 바뀌어도 디자인변경으로
  재캡처돼 테두리 없는 캡처만 쌓였다. 빈 문자열이면 제외 안 함.
- `CAPTURE_ON_HIT` / `CAPTURE_FULL_PAGE` — 캡처 on/off, 전체 페이지 여부.
- 캡처 중복 방지 — 아래 "5. 캡처 규칙" 참고 (`CAPTURE_STATE_NAME`, `FINGERPRINT_*`, `HIGHLIGHT_*`).
- 캡처 표시 — `CAPTURE_LEGEND`(맨 위 범례 박스) / `CAPTURE_BADGE`(번호 배지) / `CAPTURE_REVEAL_HIDDEN`(접힌 메뉴 펼쳐 캡처)
  on/off, `REVEAL_HOVER_WAIT_MS`·`REVEAL_MAX_CAPTURES` 등. 아래 "5. 캡처 규칙 › 캡처에서 hit 요소 찾기" 참고.
- `COOKIE_CONSENT_SELECTORS` — 쿠키 배너 동의 버튼 셀렉터 목록. 사이트마다 배너가 달라(TrustArc / 자체 `cookie-bar`)
  여러 개를 두고, 먼저 보이는 것 하나를 누른다. 새 종류의 배너가 캡처를 가리면 여기에 한 줄 추가한다.
- `CAPTURE_HIDE_SELECTOR` — 캡처 직전에 숨길 요소. 동의 버튼을 못 눌러 배너가 남아도 hit 요소를 가리지 않게 한다.
  쿠키 배너는 감지 대상이 아니다 — 배너가 hit 요소 위를 덮으면 번호 배지가 배너에 붙은 것처럼 보일 뿐이다.
- `MAX_WORKERS` — 동시 브라우저 수 (기본 4).
- `RETRY_COUNT` / `RETRY_HTTP_STATUS` / `RETRY_DELAY_SEC` — 타임아웃·브라우저 오류와 HTTP 400·429·5xx 는 **새 브라우저 컨텍스트로** 쉬었다가 다시 연다(기본 2회). 404·리다이렉트는 재시도 없이 바로 `접근실패`.
  작업(사이트×페이지)마다 컨텍스트를 새로 만들어 쿠키·세션 상태가 다음 사이트로 넘어가지 않게 한다.
- `OUTPUT_DIR` — 기본 `스크립트 폴더\output` (리포트·기록·로그).
- `CAPTURE_DIR` — 캡처 저장 폴더(절대경로). 작업용·운영 폴더가 같은 `.py` 를 쓰므로 **작업용 폴더에서 테스트해도 캡처는 여기에 쌓인다.**

## 3. 실행

```bash
python promo_link_checker_v1.2.py                    # SITECODES 전체
python promo_link_checker_v1.2.py --sitecodes ae,hq # 일부 국가만
python promo_link_checker_v1.2.py --pages home       # 일부 페이지만
python promo_link_checker_v1.2.py --no-capture       # hit 여도 캡처 안 함
python promo_link_checker_v1.2.py --debug            # 사이트별 클릭 요소 수 등 진단 로그
```

## 4. 출력

```
promo_link_checker/
├── promo_link_checker_v1.2.py
├── README.md
└── output/
    ├── _daily_report.xlsx            ← 일일 누적 리포트
    ├── _run_latest.log               ← 매 실행 덮어씀 (스케줄 실행 추적용)
    └── _capture_state.json           ← 이미 캡처한 요소 기록 (지우면 처음부터 다시 캡처)

<CAPTURE_DIR>/                        ← 캡처 저장 폴더 (절대경로 상수, output 과 별개 위치)
└── <sitecode>/
    ├── <sitecode>_<page>_<YYMMDD>_<HHMM>.png       ← 화면에 보이는 새 요소·디자인 변경이 있을 때만 (전체 페이지)
    └── <sitecode>_<page>_menu_<YYMMDD>_<HHMM>.png  ← 새 hit 요소가 접힌 메뉴 안에 있을 때, 메뉴를 펼친 화면
```

캡처는 `CAPTURE_DIR`(절대경로) 아래 **sitecode 폴더**에 모은다 (v1.2, 날짜 폴더 없음). 리포트·기록·로그는
`output\` 에 그대로 남는다. 요소당 1회만 찍으므로 폴더 안에는 새 요소·디자인 변경이
있었던 날의 캡처만 쌓이고, 찍은 날짜·시각은 파일명 맨 뒤(`_<YYMMDD>_<HHMM>`, `CAPTURE_TIME_FORMAT`)에 남는다.
`CAPTURE_DIR` 이 `output\` 밖이면 리포트·기록에는 캡처가 절대경로로 적힌다.
아래에서 `_menu.png` 라고 부르는 것은 이 `…_menu_<YYMMDD>_<HHMM>.png` 파일이다.

### _daily_report.xlsx
| A | B | C | D | E | F | G … |
|---|---|---|---|---|---|---|
| sitecode | page | url | **오늘** 결과 | 매칭키워드 | 매칭링크 | 어제 결과 … |

- 행 = (sitecode, page). 한 국가가 home·offer 2행.
- **최신 날짜 블록이 항상 D~F열**이고, 과거 날짜는 오른쪽으로 밀린다. 열어서 A~F 만 보면 오늘 상태다.
- 결과 값: `O`(녹색, hit) / `X`(없음) / `접근실패(HTTP 404·redirect·timeout·error)`(빨강) / `(미실행)`.
- `O` 셀은 **캡처 PNG 로 하이퍼링크**된다 — 그날 새로 찍었으면 새 캡처, 아니면 그 요소의 기존 캡처.
  새 요소가 전부 접힌 메뉴 안이면 전체 캡처 대신 **메뉴를 펼친 캡처**(`_menu.png`)로 연결된다.
  GNB/Footer 요소는 sitecode 당 한 번만 찍으므로, offer 행의 `O` 가 home 에서 찍은 캡처로 연결될 수 있다.
- 매칭링크 셀: `URL (링크 텍스트) [GNB/Footer] [숨김] 🆕신규 / 🔄디자인변경 / ♻재캡처` + `📷 새 캡처: …`(+ `📷 추가 캡처: …`)
  또는 `📷 기존 캡처: …`. `[숨김]` = 접힌 메뉴·안 보이는 슬라이드 안이라 화면에 안 보이는 요소.
- 같은 날 다시 돌리면 그 날짜 블록을 덮어쓴다(열이 늘지 않음). `--sitecodes`/`--pages` 로 일부만 돌리면
  대상 밖 행의 그날 값은 보존된다.
- 저장은 임시파일 → 무결성 확인 → 교체 방식이라 저장 중 실패해도 누적 이력이 안 날아간다.
  기존 파일이 손상돼 있으면 `.broken_<ts>` 로 치워두고 새로 만든다.

## 5. 캡처 규칙 (1회성)

- 요소 식별 = **sitecode + page + 링크 URL**(host+path, 쿼리·# 제외).
  단 **GNB/Footer 안 요소는 page 를 빼고 sitecode + 링크 URL** 로 본다 (v1.2) — 모든 페이지에 똑같이 나오므로
  home 에서 찍었으면 offer 에서는 다시 찍지 않는다. 그래서 한 sitecode 의 페이지들은 같은 워커가 순서대로 돈다.
- 요소마다 **디자인 지문**(12자리 해시)을 만든다. 재료는 `FINGERPRINT_FIELDS`:

  | 항목 | 내용 |
  |---|---|
  | `size` | 요소 크기 (`FINGERPRINT_SIZE_STEP`=20px 단위 반올림) |
  | `classes` | class 목록 — `active`·`loaded`·`swiper-slide-*` 같은 상태 class(`FINGERPRINT_CLASS_IGNORE`)는 제외 |
  | `images` | 안쪽 `<img>` src · 배경 이미지 URL (쿼리 제외, data: placeholder 제외) |
  | `text` | 요소 텍스트 — **숫자는 지운다** (가격·날짜만 바뀐 건 디자인 변경으로 안 본다) |

- 판정:
  - 처음 보는 링크 → **🆕신규**, 캡처
  - 본 적 있는 링크인데 지문이 다름 → **🔄디자인변경**, 캡처
  - 링크·지문 모두 본 적 있고 **그 캡처 파일이 실제로 있음** → 캡처 생략 (리포트는 기존 캡처로 링크)
  - 링크·지문은 본 적 있는데 캡처 파일이 없음(지워짐·옮겨짐) → **♻재캡처**, 다시 찍고 기록을 새 캡처로 바꾼다
    (`VERIFY_CAPTURE_EXISTS`, 기본 `True`). 잘못 찍힌 캡처를 다시 받으려면 **그 PNG 를 지우면** 다음 실행에서 다시 찍힌다.
- 한 페이지에 **화면에 보이는** 신규/변경 요소가 하나라도 있으면 그 페이지를 1장 찍고, 그 화면의 hit 요소에는 전부
  **빨간 실선** 테두리를 그린다. 신규/변경/기존 구분은 색이 아니라 **범례의 표기**로 한다
  (색으로도 나누려면 `HIGHLIGHT_KNOWN_STYLE`·`BADGE_COLOR_KNOWN` 값을 바꾼다).
- 신규/변경 요소가 **전부 접힌 메뉴 안**이면 전체 캡처는 생략하고 메뉴를 펼친 `_menu.png` 만 남긴다 (v1.2).
  전체 캡처에는 테두리가 하나도 없어 볼 게 없기 때문. 메뉴를 하나도 못 펼쳤을 때만 범례를 남기려고 전체 캡처를 찍는다.
- **캡처에서 hit 요소 찾기** (v1.1):
  - 캡처 맨 위 **범례 박스**에 hit 요소를 한 줄씩 적는다 — `번호 [본문|GNB/Footer] 링크 텍스트 → URL (신규/디자인변경/기존)`.
  - 화면에 보이는 요소는 테두리 왼쪽 위에 범례와 같은 **번호 배지**가 붙는다.
  - 접힌 GNB 메뉴 안처럼 **화면에 안 보이는 요소**는 전체 캡처에 테두리가 없다(범례에 `⚠ 화면에 안 보임` 표기).
    이때는 그 요소의 가장 가까운 보이는 조상(GNB 1단계 메뉴 등)에 마우스를 올려 메뉴를 펼친 뒤
    **현재 화면을 `_menu.png` 로** 찍는다. 마우스를 올려도 안 나타나면(캐러셀의 안 보이는 슬라이드 등) 건너뛴다.
    이미 찍은 적 있는 숨은 요소 때문에 메뉴를 다시 펼쳐 찍지는 않는다.
  - 순서는 펼친 캡처 → 범례 삽입 → 전체 캡처. 범례를 먼저 넣으면 GNB 메뉴가 hover 로 열리지 않는다.
- 기록은 `output/_capture_state.json`. **캡처에 성공했을 때만** 기록한다(실패하면 다음 실행에서 다시 시도).
  `--no-capture` 로 돌린 날도 기록하지 않는다. 파일을 지우면 모든 요소를 처음부터 다시 캡처한다.
- 같은 요소가 매일 다시 찍힌다면 지문이 불안정한 것 — `FINGERPRINT_FIELDS` 에서 항목을 빼거나
  `FINGERPRINT_CLASS_IGNORE` 에 그 class 를 추가한다.

## 6. 매일 자동 실행 (작업 스케줄러)

**작업용 폴더와 운영 폴더를 나눈다.**

| 폴더 | 용도 |
|---|---|
| `promo_link_checker\` | 작업용 — 코드 수정·테스트. 스케줄러가 돌리지 않는다 |
| `promo_link_checker_<캠페인>\` | 운영 사본 (캠페인별 KEYWORDS) — **스케줄러가 이 폴더의 코드를 돌린다**. 리포트·캡처·기록도 여기에 쌓인다 |

작업용에서 코드를 고친 뒤 운영에 반영하려면 **`.py` 만** 운영 폴더에 덮어쓴다 (`output\` 은 건드리지 않음).
다른 캠페인을 추가로 감시하려면 운영 폴더를 하나 더 복사해 `KEYWORDS` 만 바꾸고 작업명을 달리 등록한다.

작업명 `promo_link_checker_<캠페인>`, 매일 **09:31 · 12:01**, `pythonw.exe` 직접 실행(콘솔 창 없음):

```powershell
$script   = "<운영 폴더>\promo_link_checker_v1.2.py"
$action   = New-ScheduledTaskAction -Execute "C:\Python314\pythonw.exe" -Argument "`"$script`"" -WorkingDirectory (Split-Path $script)
$trigger  = @("09:31", "12:01") | ForEach-Object { $t = New-ScheduledTaskTrigger -Daily -At $_; $t.EndBoundary = "2026-12-10T23:59:59"; $t }
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName "promo_link_checker_<캠페인>" -Action $action -Trigger $trigger -Settings $settings -Force

Start-ScheduledTask -TaskName "promo_link_checker_<캠페인>"                             # 즉시 1회 실행
Get-ScheduledTask -TaskName "promo_link_checker_<캠페인>" | Get-ScheduledTaskInfo         # 다음 실행 / 마지막 결과
```

- 트리거는 하루 2번(09:31, 12:01)이고 둘 다 **종료일**(`EndBoundary`)이 걸려 있어 그 뒤로는 실행되지 않는다.
  같은 날 두 번째 실행은 그 날짜 블록을 덮어쓴다.
- `StartWhenAvailable` — 예약 시각에 PC 가 꺼져 있었으면 켜진 뒤 바로 실행한다.
- `ExecutionTimeLimit 2h` / `MultipleInstances IgnoreNew` — 멈춘 실행을 끊고, 겹쳐 돌지 않게 한다.
- 로그온 상태에서만 실행된다(기본 계정 설정).
- 경로는 전부 스크립트 폴더 기준이라 작업 디렉터리에 의존하지 않는다.
- 콘솔 출력이 없으니 진행 상황은 `output\_run_latest.log` 로 본다.

## 7. 알려진 제약

- **한국 IP 기준** 결과다. 국가별 IP 리다이렉션이 걸리는 사이트(예: 운영 중단된 `ge`)는 열리지 않는다.
- 스크롤로 레이지 로딩 영역을 띄운 뒤 검사하지만, 탭·캐러셀 **안 보이는 슬라이드**의 링크는 DOM 에
  있으면 잡히고, 클릭해야 불러오는 콘텐츠는 못 잡는다.
- KV 텍스트·CTA 는 화면 안에 잠깐 머물러야 페이드인되는데, 빠른 스크롤이 건너뛰면 `[숨김]` 판정돼 전체 캡처에
  테두리가 빠진다. 그래서 숨김 판정된 본문 hit 요소는 하나씩 화면 가운데로 스크롤해 `RESCROLL_HIDDEN_WAIT_MS`(기본 1.2초)
  기다린 뒤 다시 판정한다 (`RESCROLL_HIDDEN`). 크기 0 인 복제본(display:none)·GNB/Footer 요소는 제외.
- 전체 캡처는 화면 높이를 페이지 전체로 늘려 찍는데, 그 순간 KV 가 이미지를 다시 골라 저화질 미리보기가 찍힐 수 있다.
  그래서 화면 높이를 미리 늘려 두고 `CAPTURE_SETTLE_MS`(기본 1.5초) 기다렸다가 찍는다.
- GNB 메가메뉴처럼 화면에 안 보이는 링크가 hit 이면 메뉴를 펼친 `_menu.png` 에 표시된다 (전체 캡처에는 테두리가 없다).
  마우스를 올려서 열리는 메뉴만 펼칠 수 있다 — 클릭해야 열리는 메뉴·캐러셀 슬라이드는 범례와 리포트 `[숨김]` 태그로만 알 수 있다.
