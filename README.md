# promo_link_checker  
<sub>2026-11-23  Jonghyun Park w/ Claude</sub>  
국가별 **home · 프로모션(offer) 페이지의 클릭 가능한 요소** 중에 타겟 캠페인 링크가 걸려 있는지 매일 확인하고,
결과를 `_daily_report.xlsx` 에 하루 1블록씩 누적하는 도구. 캡처는 **처음 보는 hit 요소, 또는 디자인이
바뀐 요소가 있을 때만 1회** 찍는다(매일 누적 캡처하지 않음).

- 단일 파일 `promo_link_checker_v1.5.py` — 설정은 전부 상단 `사용자가 바꿔야 하는 부분` 블록.

---

## 1. 준비

- Windows / Python 3.10+ / Chrome 설치
- `pip install playwright openpyxl`
- 메일 알림(`MAIL_ON_FIRST_HIT`)을 켤 때만: `pip install pywin32` + Outlook 데스크톱 앱
- 영역 이미지 해시(`IMAGE_HASH_MODE`, 기본 `log`)용: `pip install pillow`. 없으면 경고만 남기고 해시를 건너뛴다
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
- 줄 끝에 **`@sitecode,sitecode`** 를 붙이면 그 sitecode 페이지에서만 쓰는 규칙이 된다 (없으면 전 sitecode).
  흔한 표현(예: `gift`)은 다른 나라 상시 페이지(GNB 메뉴 등)에도 걸리므로, 실제로 캠페인에 쓴 sitecode 로 한정한다.
  리포트·메일의 매칭키워드에는 `@…` 를 뺀 키워드만 적힌다.

  ```python
  KEYWORDS = """
  summer sale
  galaxy gifts @pl,bg
  """
  ```
  → `…/uk/mobile/galaxygifts/` 는 uk 가 목록에 없어 ❌, `…/pl/offer/galaxy-gifts/` 는 ✅.

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

### 메일 알림 — 처음 hit 된 sitecode (v1.4)
**한 번도 hit 된 적 없던 sitecode** 에서 타겟 링크가 처음 발견되면 메일로 알린다. 아직 캠페인을 시작하지 않은
국가가 언제 시작했는지 리포트를 매일 열어 보지 않아도 알 수 있다.

| 상수 | 의미 |
|---|---|
| `MAIL_ON_FIRST_HIT` | 알림 on/off |
| `MAIL_TO` | 받는 사람(여러 명이면 세미콜론). 비우면 Outlook 기본 계정 = 자기 자신 |
| `MAIL_SUBJECT_PREFIX` | 제목 맨 앞에 붙는 말. 폴더 이동 대상도 이 접두로 찾는다 |
| `MAIL_MOVE_TO_FOLDER` | 받은 알림 메일을 옮길 폴더 — 받은편함 기준 하위 경로(`/` 구분). 없으면 만든다. 비우면 안 옮긴다 |
| `MAIL_MOVE_WAIT_SEC` | 보낸 메일이 받은편함에 도착하기를 기다리는 상한(초) |
| `MAIL_INLINE_IMAGE_WIDTH` / `MAIL_MAX_IMAGES_PER_SITE` | 본문 이미지 가로폭 / sitecode 당 넣을 캡처 장수 |

- **"첫 hit" 판정** = 실행 시작 시점의 캡처 기록(`_capture_state.json`)에 그 sitecode 가 **아예 없는데** 이번 실행에서 hit.
  home · offer 중 어느 쪽이든 한 번이라도 기록된 sitecode 는 다시 알리지 않는다. 한 실행에서 여러 sitecode 가 걸리면 메일 1통에 묶는다.
- 메일 내용 = sitecode 별 페이지 URL · 매칭 키워드 · 매칭 링크 목록 · 캡처 이미지(**본문에 바로 보이고**, 원본도 첨부로 붙는다).
- 본문에는 보내는 PC 의 절대경로를 적지 않는다. 캡처 파일은 `CAPTURE_DIR` 기준 상대경로(`sitecode\파일명`)로, 캡처 폴더는 맨 아래에 한 번만 적는다
  (사용자 홈 아래면 홈을 뗀 경로). 받는 사람마다 로컬 경로 앞부분이 달라서다.
- 로컬 Outlook 앱으로 보낸다(별도 인증 없음). 보낸 직후 받은편함에 도착하면 `MAIL_MOVE_TO_FOLDER` 로 옮기고,
  제때 안 오면 다음 실행 때 옮긴다.
- **캡처에 성공한 hit 만 알린다.** 캡처를 끄면(`CAPTURE_ON_HIT=False` · `--no-capture`) 기록이 안 남아 알림도 가지 않는다.
- 메일 발송이 실패해도 경고 로그만 남고 검사 · 리포트에는 영향이 없다. 단 기록은 이미 남으므로 그 sitecode 는 다시 알리지 않는다
  (`--no-mail` 로 돌린 실행에서 처음 hit 된 경우도 같다) — 로그의 `새로 캡처` 줄로 확인한다.

### 그 외
- `EXCLUDE_GLOBAL_UI` — `True` 면 GNB/Footer 안 링크는 hit 에서 뺀다. 기본 `False`(페이지 전체)이고,
  리포트에는 `[GNB/Footer]` 태그가 붙어 구분된다.
- `EXCLUDE_HIT_SELECTOR` — 이 선택자 안(또는 자신)의 링크는 키워드가 맞아도 hit 에서 뺀다. 기본값은 채팅봇 위젯
  프로모 배너(`[class*='rcw-promo']`) — 닫힌 위젯 안이라 화면에 안 보이는데, 배너 이미지만 바뀌어도 디자인변경으로
  재캡처돼 테두리 없는 캡처만 쌓였다. Insider 개인화 팝업 링크(`.ins-element-link`, `.ins-content-link`)도
  같은 이유로 제외한다 — 0×0 레이어인데 디자인변경으로 재캡처됐다. 빈 문자열이면 제외 안 함.
- `CAPTURE_ON_HIT` / `CAPTURE_FULL_PAGE` — 캡처 on/off, 전체 페이지 여부.
- `CAPTURE_IMAGE_EXT` / `CAPTURE_JPEG_QUALITY` — 캡처 형식(기본 `.jpg`, 품질 82). `.png` 로 바꾸면 무손실로 저장한다.
- 캡처 중복 방지 — 아래 "5. 캡처 규칙" 참고 (`CAPTURE_STATE_NAME`, `FINGERPRINT_*`, `HIGHLIGHT_*`).
- 영역 이미지 해시 — 아래 "5. 캡처 규칙 › 영역 이미지 해시" 참고 (`IMAGE_HASH_MODE`, `IMAGE_HASH_THRESHOLD` 등).
- 캡처 표시 — `CAPTURE_LEGEND`(맨 위 범례 박스) / `CAPTURE_BADGE`(번호 배지) / `CAPTURE_REVEAL_HIDDEN`(접힌 메뉴 펼쳐 캡처) /
  `CAPTURE_REVEAL_TABS`(안 열린 탭 눌러 캡처) on/off, `REVEAL_HOVER_WAIT_MS`·`REVEAL_TAB_WAIT_MS`·`REVEAL_MAX_CAPTURES` 등. 아래 "5. 캡처 규칙 › 캡처에서 hit 요소 찾기" 참고.
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
python promo_link_checker_v1.5.py                    # SITECODES 전체
python promo_link_checker_v1.5.py --sitecodes ae,hq # 일부 국가만
python promo_link_checker_v1.5.py --pages home       # 일부 페이지만
python promo_link_checker_v1.5.py --no-capture       # hit 여도 캡처 안 함
python promo_link_checker_v1.5.py --no-mail          # 처음 hit 된 sitecode 가 있어도 메일 안 보냄
python promo_link_checker_v1.5.py --debug            # 사이트별 클릭 요소 수 등 진단 로그
```

## 4. 출력

```
promo_link_checker/
├── promo_link_checker_v1.5.py
├── README.md
└── output/
    ├── _daily_report.xlsx            ← 일일 누적 리포트
    ├── _run_latest.log               ← 매 실행 덮어씀 (스케줄 실행 추적용)
    ├── _capture_state.json           ← 이미 캡처한 요소 기록 (지우면 처음부터 다시 캡처)
    ├── _image_hash_state.json        ← 요소 영역별 마지막 이미지 해시 (IMAGE_HASH_MODE 가 off 가 아닐 때)
    └── _image_hash_log.csv           ← 실행마다 요소 영역별 해시 거리 누적 (기준값 정하기용)

<CAPTURE_DIR>/                        ← 캡처 저장 폴더 (절대경로 상수, output 과 별개 위치)
└── <sitecode>/
    ├── <sitecode>_<page>_<YYMMDD>_<HHMM>.jpg       ← 화면에 보이는 새 요소·디자인 변경이 있을 때만 (전체 페이지)
    ├── <sitecode>_<page>_menu_<YYMMDD>_<HHMM>.jpg  ← 새 hit 요소가 접힌 메뉴 안에 있을 때, 메뉴를 펼친 화면
    └── <sitecode>_<page>_tab_<YYMMDD>_<HHMM>.jpg   ← 새 hit 요소가 안 열린 탭 안에 있을 때, 탭을 눌러 연 화면
```

캡처는 `CAPTURE_DIR`(절대경로) 아래 **sitecode 폴더**에 모은다 (v1.2, 날짜 폴더 없음). 리포트·기록·로그는
`output\` 에 그대로 남는다. 요소당 1회만 찍으므로 폴더 안에는 새 요소·디자인 변경이
있었던 날의 캡처만 쌓이고, 찍은 날짜·시각은 파일명 맨 뒤(`_<YYMMDD>_<HHMM>`, `CAPTURE_TIME_FORMAT`)에 남는다.
`CAPTURE_DIR` 이 `output\` 밖이면 리포트·기록에는 캡처가 절대경로로 적힌다.
아래에서 `_menu.jpg` / `_tab.jpg` 라고 부르는 것은 이 `…_menu_<YYMMDD>_<HHMM>.jpg` / `…_tab_<YYMMDD>_<HHMM>.jpg` 파일이다.

### _daily_report.xlsx
| A | B | C | D | E | F | G … |
|---|---|---|---|---|---|---|
| sitecode | page | url | **오늘** 결과 | 매칭키워드 | 매칭링크 | 어제 결과 … |

- 행 = (sitecode, page). 한 국가가 home·offer 2행.
- **최신 날짜 블록이 항상 D~F열**이고, 과거 날짜는 오른쪽으로 밀린다. 열어서 A~F 만 보면 오늘 상태다.
- 결과 값: `O`(녹색, hit) / `X`(없음) / `접근실패(HTTP 404·redirect·timeout·error)`(빨강) / `(미실행)`.
- `O` 셀은 **캡처 이미지로 하이퍼링크**된다 — 그날 새로 찍었으면 새 캡처, 아니면 그 요소의 기존 캡처.
  새 요소가 전부 접힌 메뉴·안 열린 탭 안이면 전체 캡처 대신 **펼친 캡처**(`_menu.jpg` / `_tab.jpg`)로 연결된다.
  GNB/Footer 요소는 sitecode 당 한 번만 찍으므로, offer 행의 `O` 가 home 에서 찍은 캡처로 연결될 수 있다.
- 매칭링크 셀: `URL (링크 텍스트) [GNB/Footer] [숨김] 🆕신규 / 🔄디자인변경 / ♻재캡처` + `📷 새 캡처: …`(+ `📷 추가 캡처: …`)
  또는 `📷 기존 캡처: …`. `[숨김]` = 접힌 메뉴·안 열린 탭·안 보이는 슬라이드 안이라 화면에 안 보이는 요소.
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
  | `size` | 요소 크기 (`FINGERPRINT_SIZE_STEP`=20px 단위 반올림). 비교할 땐 허용오차가 있다 — 아래 판정 참고 |
  | `classes` | class 목록 — `active`·`loaded`·`swiper-slide-*` 같은 상태 class(`FINGERPRINT_CLASS_IGNORE`)는 제외 |
  | `images` | 안쪽 `<img>` src · 배경 이미지 URL (쿼리 제외, data: placeholder 제외) |
  | `text` | 요소 텍스트 — **숫자는 지운다** (가격·날짜만 바뀐 건 디자인 변경으로 안 본다) |

- 판정:
  - 처음 보는 링크 → **🆕신규**, 캡처
  - 본 적 있는 링크인데 지문이 다름 → **🔄디자인변경**, 캡처
    단 class·이미지·텍스트가 같고 크기 차이가 가로·세로 각각 `FINGERPRINT_SIZE_TOLERANCE`(20px) 이하이거나,
    어느 한쪽이 0×0(안 열린 탭·캐러셀 안 숨김 상태)이면 **같은 디자인**으로 본다 —
    반올림 경계(예: 90px)에서 1~2px 흔들려 80 ↔ 100 으로 갈리던 재캡처 방지
  - 링크·지문 모두 본 적 있고 **그 캡처 파일이 실제로 있음** → 캡처 생략 (리포트는 기존 캡처로 링크)
  - 링크·지문은 본 적 있는데 캡처 파일이 없음(지워짐·옮겨짐) → **♻재캡처**, 다시 찍고 기록을 새 캡처로 바꾼다
    (`VERIFY_CAPTURE_EXISTS`, 기본 `True`). 잘못 찍힌 캡처를 다시 받으려면 **그 캡처 파일을 지우면** 다음 실행에서 다시 찍힌다.
- 한 페이지에 **화면에 보이는** 신규/변경 요소가 하나라도 있으면 그 페이지를 1장 찍고, 그 화면의 hit 요소에는 전부
  **빨간 실선** 테두리를 그린다. 신규/변경/기존 구분은 색이 아니라 **범례의 표기**로 한다
  (색으로도 나누려면 `HIGHLIGHT_KNOWN_STYLE`·`BADGE_COLOR_KNOWN` 값을 바꾼다).
- 신규/변경 요소가 **전부 접힌 메뉴·안 열린 탭 안**이면 전체 캡처는 생략하고 펼친 `_menu.jpg` / `_tab.jpg` 만 남긴다 (v1.2).
  전체 캡처에는 테두리가 하나도 없어 볼 게 없기 때문. 메뉴를 하나도 못 펼쳤을 때만 범례를 남기려고 전체 캡처를 찍는다.
- **캡처에서 hit 요소 찾기** (v1.1):
  - 캡처 맨 위 **범례 박스**에 hit 요소를 한 줄씩 적는다 — `번호 [본문|GNB/Footer] 링크 텍스트 → URL (신규/디자인변경/기존)`.
  - 화면에 보이는 요소는 테두리 왼쪽 위에 범례와 같은 **번호 배지**가 붙는다.
  - 접힌 GNB 메뉴 안처럼 **화면에 안 보이는 요소**는 전체 캡처에 테두리가 없다(범례에 `⚠ 화면에 안 보임` 표기).
    이때는 그 요소의 가장 가까운 보이는 조상(GNB 1단계 메뉴 등)에 마우스를 올려 메뉴를 펼친 뒤
    **현재 화면을 `_menu.jpg` 로** 찍는다. 마우스를 올려도 안 나타나면(캐러셀의 안 보이는 슬라이드 등) 건너뛴다.
    이미 찍은 적 있는 숨은 요소 때문에 메뉴를 다시 펼쳐 찍지는 않는다.
  - **안 열린 탭 안 요소** (`CAPTURE_REVEAL_TABS`): 본문 요소는 hover 전에 먼저 그 요소를 품은 탭 패널의
    탭 버튼을 눌러 연다 → 요소를 화면 가운데로 스크롤 → 테두리·배지 → **`_tab.jpg` 로** 찍는다.
    찍는 범위 = 테두리 친 요소들을 감싼 영역 + 위아래 `REVEAL_TAB_CONTEXT_PX`(기본 700px) — 앞뒤 콘텐츠가 뭔지 보이게.
    그 바로 위에 전체 캡처와 같은 **범례 박스**(`… · 탭 '<탭명>' 을 눌러 연 화면`, 번호별 링크)를 붙인다.
    탭은 `aria-controls` / `href="#패널id"` / `data-target` / `aria-labelledby` 로 찾고, 없으면 같은 개수의 탭 버튼 목록에서
    패널 순번으로 고른다. 페이지 이동을 막으려고 button · `[role=tab]` · href 가 `#`/`javascript:`/없음 인 `<a>` 만 누른다.
    찍는 순간에만 화면에 떠 있는(`position: fixed`) 팝업·채팅 위젯을 숨긴다(`REVEAL_TAB_HIDE_FLOATING`).
  - 테두리는 요소 outline 과 함께 **맨 위 레이어에 같은 사각 프레임**을 한 겹 더 그린다 — 부모가 overflow 로 outline 을
    잘라내는 카드형 링크도 테두리가 보이게. 전체 캡처의 테두리는 그 시점에 실제로 보이는 요소 기준이다.
  - 순서는 펼친 캡처 → 범례 삽입 → 전체 캡처. 범례를 먼저 넣으면 GNB 메뉴가 hover 로 열리지 않는다.
- 기록은 `output/_capture_state.json`. **캡처에 성공했을 때만** 기록한다(실패하면 다음 실행에서 다시 시도).
  `--no-capture` 로 돌린 날도 기록하지 않는다. 파일을 지우면 모든 요소를 처음부터 다시 캡처한다.
- 같은 요소가 매일 다시 찍힌다면 지문이 불안정한 것 — `FINGERPRINT_FIELDS` 에서 항목을 빼거나
  `FINGERPRINT_CLASS_IGNORE` 에 그 class 를 추가한다.

### 영역 이미지 해시 (2026-11-23)

위 디자인 지문은 HTML 속성만 보므로, **이미지 URL 은 그대로인데 그림만 바뀐 배너**는 못 잡는다. 그래서
화면에 보이는 hit 요소 영역만 따로 찍어(애니메이션 정지) **dHash**(64비트 이미지 해시)를 만들고,
지난 실행의 같은 요소 해시와 **다른 비트 수(거리 0~64)** 를 잰다. Pillow 만 쓴다 (별도 이미지 비교 라이브러리 없음).

| 설정 | 기본 | 내용 |
|---|---|---|
| `IMAGE_HASH_MODE` | `log` | `off` = 안 함 / `log` = 거리만 기록 (캡처 판정에는 안 씀) / `detect` = 거리가 기준값을 넘으면 🔄디자인변경으로 다시 캡처 |
| `IMAGE_HASH_THRESHOLD` | `10` | 이 값보다 거리가 크면 "바뀜". 압축·렌더링 노이즈는 보통 0 ~ 5 |
| `IMAGE_HASH_MAX_PER_PAGE` | `20` | 한 페이지에서 영역을 찍을 요소 최대 개수 |
| `IMAGE_HASH_MIN_SIZE` | `16` | 가로·세로가 이 px 보다 작은 요소는 건너뜀 |

- 요소 식별 = `링크키#순번` (같은 링크가 한 페이지에 여러 요소면 DOM 순서로 1, 2, …). 숨은 요소(접힌 메뉴·안 보이는 슬라이드)는 찍지 않는다.
- `output/_image_hash_log.csv` 에 실행마다 `sitecode · page · 번호 · 요소 · DOM 판정 · 해시 · 이전 해시 · 거리 · 기준 초과(Y)` 를 쌓는다.
  기준을 넘은 요소는 실행 로그 끝에 `🖼 영역 이미지 거리 > 10 …` 한 줄로도 남는다.
- **운영 절차**: 처음 며칠은 `log` 로 두고 로그의 거리 분포를 본다 → 캐러셀·카운트다운처럼 매번 흔들리는 요소가 없는지 확인 →
  기준값을 정한 뒤 `detect` 로 바꾼다.

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
$script   = "<운영 폴더>\promo_link_checker_v1.5.py"
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
- GNB 메가메뉴처럼 화면에 안 보이는 링크가 hit 이면 메뉴를 펼친 `_menu.jpg` 에 표시된다 (전체 캡처에는 테두리가 없다).
  마우스를 올려서 열리는 메뉴만 펼칠 수 있다 — 클릭해야 열리는 메뉴·캐러셀 슬라이드는 범례와 리포트 `[숨김]` 태그로만 알 수 있다.

## 신규 · 디자인변경 요소 알림 메일 (`promo_change_mail.py`)
본체 메일은 sitecode 가 **처음** hit 될 때만 간다. 이미 hit 기록이 있는 sitecode 에서 새 클릭 요소가 생기거나
디자인이 바뀐 것은 캡처 · 리포트에만 남으므로, 그걸 모아 알리는 별도 도구다. **checker 와 같은 폴더에 두고 실행**한다.

- **설정 재사용**: 같은 폴더의 `promo_link_checker_v*.py`(가장 높은 버전)를 불러와 `OUTPUT_DIR` · `CAPTURE_DIR` ·
  `MAIL_SUBJECT_PREFIX` · `MAIL_MOVE_TO_FOLDER` 등을 그대로 쓴다 → 캠페인 사본마다 이 파일만 복사하면 된다.
- **대상**: `_capture_state.json` 에서 `SINCE_DAYS`(기본 1 = 오늘) 안에 처음 기록된 `KINDS`(기본 신규 · 디자인변경) 요소.
- **중복 방지**: 보낸 요소는 `output/_change_mail_sent.json` 에 남겨 하루에 몇 번 돌려도 같은 요소는 다시 안 간다.
- **본문**: 종류별 절(신규 → 디자인변경, `KINDS` 순서) 안에 sitecode 별 비교표 `페이지 | 구분 · 링크 | 전 | 후` — **전 / 후 캡처 이미지를 한 행에 나란히** 놓는다
  (이미지 가로 `COMPARE_IMAGE_WIDTH`, sitecode 당 최대 `MAIL_MAX_IMAGES_PER_SITE` 행, 넘치면 파일명 목록).
  - 맨 위에 종류별 site 요약(`신규 4개 site: jp(5), sa(6) …`). 같은 표에서 이미 나온 캡처는 다시 싣지 않고 `↑ 위 행과 같은 캡처` 로 적는다.
  - 같은 캡처 쌍을 공유하는 요소는 한 행으로 묶는다. 전 / 후 칸 위에 요소 설명(날짜 · 카드 종류 · 크기 · 이미지 파일명 / 텍스트)을 적는다.
  - 디자인변경: 같은 링크의 **앞선 날짜** 기록과 비교한다. 같은 날 같은 링크로 함께 잡힌 다른 요소(이미지 카드 · 텍스트 카드 등)는
    checker 가 디자인변경으로 적어도 신규로 본다. 신규: 전은 `— (신규)`.
- **제목**: `<MAIL_SUBJECT_PREFIX>[변경] MM-DD 신규 N · 디자인변경 N (N개 site: …)`. checker 접두를 그대로 쓰므로
  보낸 직후 받은 메일을 `MAIL_MOVE_TO_FOLDER` 로 옮긴다 (늦게 오면 다음 checker 실행 때 같이 옮겨진다).

| 상수 | 의미 |
|---|---|
| `CHANGE_MAIL_TO` | 받는 사람 (세미콜론 구분). 비우면 자기 자신. checker 의 `MAIL_TO` 와 따로 둔다 |
| `KINDS` | 알릴 종류 — `new` 신규 / `changed` 디자인변경 / `recapture` 재캡처 |
| `SKIP_FIRST_HIT_SITES` | 기간 안에 처음 hit 된 sitecode 는 뺀다 — checker 의 첫 hit 메일과 겹치지 않게 |
| `SCHEDULE_XLSX` · `SCHEDULE_SHEET` · `SCHEDULE_HEADER_ROW` · `SCHEDULE_SITE_HEADER` · `SCHEDULE_PERIOD_COLS` | 캠페인 일정 엑셀 — 오늘이 기간(시작·끝 열 쌍) 안인데 checker 오늘 결과(일일 리포트)에 hit 가 없는 site 를 맨 위 요약에 "기간 진행중 · 감지 안 됨" 으로 적는다 (offer 없음 · 점검 목록 밖 · 오늘 결과 없음 · 접속 실패는 괄호 — checker 에서 `/offer/ 없음` 사유로 주석 처리된 site 는 offer 없음). 비우면 안 적는다. 변경 메일이 나갈 때만 같이 나간다 |
| `SCHEDULE_API_EXCLUDE_HEADER` | 일정 엑셀에서 API 수기 제외를 표시한 열 머리글 (기본 `api수기제외`). 이 열이 `O` 인 site 는 "감지 안 됨" 에서 빼서 **다음 줄 `(API 제외 site)`** 에 회색으로 적고, API 제외가 아닌데 감지 안 된 site 를 먼저 굵게 적는다. 비우거나 열이 없으면 한 줄 |
| `SINCE_DAYS` | 오늘 포함 며칠 안의 기록을 볼지 |
| `CARD_KINDS` | class 문자열 → 카드 종류 이름 (요소 설명에 쓴다) |
| `MOVE_WAIT_SEC` | 보낸 메일이 받은편함에 올 때까지 기다렸다가 checker 의 `MAIL_MOVE_TO_FOLDER` 로 옮기는 최대 시간(초) |
| `ATTACH_ORIGINALS` | 원본 캡처를 첨부로도 붙일지 (`False` 면 메일 용량이 절반 정도) |

```powershell
python promo_change_mail.py              # 보낼 게 있으면 보낸다
python promo_change_mail.py --dry-run    # 대상 · 본문만 확인 (output/_change_mail_preview.html)
python promo_change_mail.py --test       # 본인에게만 [TEST] 로 (보낸 기록 안 남김)
python promo_change_mail.py --day 2026-11-20 --resend   # 그날 요소를 다시 보냄
```

checker 실행이 끝난 뒤 돌려야 그 실행의 새 캡처가 들어간다 → 작업 스케줄러에서 checker 작업의 **두 번째 동작**으로 등록한다 (동작은 순서대로 돌아 감지가 끝나면 바로 메일이 나간다). `pythonw` 로 돌면 출력은 `output/_change_mail_latest.log`.
