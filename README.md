# promo_link_checker  
<sub>2026-09-29  Jonghyun Park w/ Claude</sub>  
국가별 **home · 프로모션(offer) 페이지의 클릭 가능한 요소** 중에 타겟 캠페인 링크가 걸려 있는지 매일 확인하고,
결과를 `_daily_report.xlsx` 에 하루 1블록씩 누적하는 도구. 캡처는 **처음 보는 hit 요소, 또는 디자인이
바뀐 요소가 있을 때만 1회** 찍는다(매일 누적 캡처하지 않음).

- 단일 파일 `promo_link_checker_v1.0.py` — 설정은 전부 상단 `사용자가 바꿔야 하는 부분` 블록.

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
- `CAPTURE_ON_HIT` / `CAPTURE_FULL_PAGE` — 캡처 on/off, 전체 페이지 여부.
- 캡처 중복 방지 — 아래 "5. 캡처 규칙" 참고 (`CAPTURE_STATE_NAME`, `FINGERPRINT_*`, `HIGHLIGHT_*`).
- `MAX_WORKERS` — 동시 브라우저 수 (기본 4).
- `RETRY_COUNT` / `RETRY_HTTP_STATUS` / `RETRY_DELAY_SEC` — 타임아웃·브라우저 오류와 HTTP 400·429·5xx 는 **새 브라우저 컨텍스트로** 쉬었다가 다시 연다(기본 2회). 404·리다이렉트는 재시도 없이 바로 `접근실패`.
  작업(사이트×페이지)마다 컨텍스트를 새로 만들어 쿠키·세션 상태가 다음 사이트로 넘어가지 않게 한다.
- `OUTPUT_DIR` — 기본 `스크립트 폴더\output`.

## 3. 실행

```bash
python promo_link_checker_v1.0.py                    # SITECODES 전체
python promo_link_checker_v1.0.py --sitecodes ae,hq # 일부 국가만
python promo_link_checker_v1.0.py --pages home       # 일부 페이지만
python promo_link_checker_v1.0.py --no-capture       # hit 여도 캡처 안 함
python promo_link_checker_v1.0.py --debug            # 사이트별 클릭 요소 수 등 진단 로그
```

## 4. 출력

```
260929_promo_link_checker/
├── promo_link_checker_v1.0.py
├── README.md
└── output/
    ├── _daily_report.xlsx            ← 일일 누적 리포트
    ├── _run_latest.log               ← 매 실행 덮어씀 (스케줄 실행 추적용)
    ├── _capture_state.json           ← 이미 캡처한 요소 기록 (지우면 처음부터 다시 캡처)
    └── capture/<MMDD>/<sitecode>_<page>_<HHMM>.png  ← 새 요소·디자인 변경이 있을 때만
```

### _daily_report.xlsx
| A | B | C | D | E | F | G … |
|---|---|---|---|---|---|---|
| sitecode | page | url | **오늘** 결과 | 매칭키워드 | 매칭링크 | 어제 결과 … |

- 행 = (sitecode, page). 한 국가가 home·offer 2행.
- **최신 날짜 블록이 항상 D~F열**이고, 과거 날짜는 오른쪽으로 밀린다. 열어서 A~F 만 보면 오늘 상태다.
- 결과 값: `O`(녹색, hit) / `X`(없음) / `접근실패(HTTP 404·redirect·timeout·error)`(빨강) / `(미실행)`.
- `O` 셀은 **캡처 PNG 로 하이퍼링크**된다 — 그날 새로 찍었으면 새 캡처, 아니면 그 요소의 기존 캡처.
- 매칭링크 셀: `URL (링크 텍스트) [GNB/Footer] 🆕신규 / 🔄디자인변경` + `📷 새 캡처: …` 또는 `📷 기존 캡처: …`.
- 같은 날 다시 돌리면 그 날짜 블록을 덮어쓴다(열이 늘지 않음). `--sitecodes`/`--pages` 로 일부만 돌리면
  대상 밖 행의 그날 값은 보존된다.
- 저장은 임시파일 → 무결성 확인 → 교체 방식이라 저장 중 실패해도 누적 이력이 안 날아간다.
  기존 파일이 손상돼 있으면 `.broken_<ts>` 로 치워두고 새로 만든다.

## 5. 캡처 규칙 (1회성)

- 요소 식별 = **sitecode + page + 링크 URL**(host+path, 쿼리·# 제외).
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
  - 링크·지문 모두 본 적 있음 → 캡처 생략 (리포트는 기존 캡처로 링크)
- 한 페이지에 신규/변경 요소가 하나라도 있으면 그 페이지를 1장 찍고, 신규/변경 요소는 **빨간 실선**,
  이미 찍은 요소는 **주황 점선** 테두리로 구분한다.
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

작업명 `promo_link_checker_<캠페인>`, 매일 **12:01**, `pythonw.exe` 직접 실행(콘솔 창 없음):

```powershell
$script   = "<운영 폴더>\promo_link_checker_v1.0.py"
$action   = New-ScheduledTaskAction -Execute "C:\Python314\pythonw.exe" -Argument "`"$script`"" -WorkingDirectory (Split-Path $script)
$trigger  = New-ScheduledTaskTrigger -Daily -At "12:01"
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName "promo_link_checker_<캠페인>" -Action $action -Trigger $trigger -Settings $settings -Force

Start-ScheduledTask -TaskName "promo_link_checker_<캠페인>"                             # 즉시 1회 실행
Get-ScheduledTask -TaskName "promo_link_checker_<캠페인>" | Get-ScheduledTaskInfo         # 다음 실행 / 마지막 결과
```

- `StartWhenAvailable` — 12:01 에 PC 가 꺼져 있었으면 켜진 뒤 바로 실행한다.
- `ExecutionTimeLimit 2h` / `MultipleInstances IgnoreNew` — 멈춘 실행을 끊고, 겹쳐 돌지 않게 한다.
- 로그온 상태에서만 실행된다(기본 계정 설정).
- 경로는 전부 스크립트 폴더 기준이라 작업 디렉터리에 의존하지 않는다.
- 콘솔 출력이 없으니 진행 상황은 `output\_run_latest.log` 로 본다.

## 7. 알려진 제약

- **한국 IP 기준** 결과다. 국가별 IP 리다이렉션이 걸리는 사이트(예: 운영 중단된 `ge`)는 열리지 않는다.
- 스크롤로 레이지 로딩 영역을 띄운 뒤 검사하지만, 탭·캐러셀 **안 보이는 슬라이드**의 링크는 DOM 에
  있으면 잡히고, 클릭해야 불러오는 콘텐츠는 못 잡는다.
- GNB 메가메뉴처럼 화면에 안 보이는 링크가 hit 이면 캡처에 테두리가 보이지 않을 수 있다(리포트 `[GNB/Footer]` 태그 참고).
