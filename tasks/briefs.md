# Briefs

리서처가 적고 나머지가 꺼내 쓴다. 출처·검증 상태를 같이 적는다.

## 2026-09-25 — 시작 전 데이터 확인
- 기존 블로그 성과 데이터 확인 결과: **쓸 수 있는 숫자 없음.**
  - blog(ordinarydayinkorea): tasks/search-stats.md 없음 (Search Console 미연결).
  - half-handy: tasks/search-stats.md 없음. 2026-09-24 분석도 "트래픽 수치 데이터 없음".
- 그래서 첫 4개 주제는 사용자가 정한 방향을 그대로 쓴다. Search Console 연결 후
  4주 뒤 analyst 가 다시 본다.
- half-handy 분석에서 가져올 교훈: 새 blogspot 에 하루 5개는 scaled content 위험 →
  Tidy Tabs 는 하루 1개. search_description 은 API로 안 먹을 수 있음 → Blogger 설정에서
  "검색 설명" 켜기 + 글별 수동 확인.

## 2026-09-25 — 직접 확인한 결과 (LibreOffice Calc 24.2.7, Linux)
- ZIP CSV 가져오기: `tidy-tabs-zip-codes-sample.csv` 를 Calc 로 열면 ZIP 열이 숫자로
  바뀌어 02108 → 2108, 00501 → 501. ZIP+4(하이픈 포함)는 텍스트로 유지.
- `=TEXT(C2,"00000")` 로 2108 → "02108", 501 → "00501" 복구 확인.
- 인쇄: 48행 x 14열 가격표, 기본 설정(세로, 여백 기본)으로 PDF 출력 시 6쪽 →
  가로 + 1페이지 맞춤 + 여백 0.25/0.5in 적용 시 US Letter 1쪽.
- 재고/프로젝트 템플릿 수식 재계산 결과가 기대값과 일치 (On hand, REORDER, Amount,
  Payment 상태, Summary 합계).
- Excel·Google Sheets 에서는 아직 직접 확인 안 함 → 글에서는 Microsoft/Google 도움말
  링크로 근거를 댄다. 사용자가 실제 화면을 찍으면 tasks/screenshots-needed.md 참고.

## 2026-09-25 — City, ST ZIP 나누기 (LibreOffice Calc 24.2.7)
- LEFT/FIND, MID/FIND+2, RIGHT(,5) 로 8개 주소 모두 정확히 분리 (Salt Lake City, St. Louis 포함).
  ZIP 은 텍스트로 반환되어 00501 유지.
- 예외: 쉼표 없음 → City/State #VALUE!, ZIP 정상. "Boston,MA" → State "A ".
  끝 공백 → ZIP "2108 ", RIGHT(TRIM()) 로 "02108". ZIP+4 → RIGHT(,5)="-1522",
  RIGHT(,10)="02108-1522", LEFT(RIGHT(,10),5)="02108".

## 2026-09-28 — 신규 주제 5개 리서치 + 템플릿 제작/검증 (LibreOffice Calc 24.2.7, Linux)
- tasks/search-stats.md 없음 (Search Console 미연결) — 숫자 기반 판단 불가, 3-gate 기준으로만 후보 선정.
- posts/ready + posts/published 제목 전수 확인, tasks/topics-seed.md 는 전부 소진됨 확인. 겹치지 않는
  신규 주제 5개 선정 (제목/슬러그/훅은 tasks/team-topics-2026-09-29.md 참고):
  1. dependent-dropdown-list-excel-google-sheets — Category > Subcategory 종속 드롭다운 (INDIRECT + named range)
  2. client-contact-list-template-google-sheets — 팔로업 알림 CRM류 시트 (TODAY() 기반 상태)
  3. packing-slip-template-google-sheets — Etsy 배송용 포장 명세서, US Letter 1쪽 인쇄
  4. running-balance-column-excel-google-sheets — 현금출납부 누적 잔액 (전행 참조 수식)
  5. highlight-duplicate-values-excel-google-sheets — 조건부 서식 COUNTIF 로 중복 값 강조 (행 삭제 아님, "중복 행 삭제" 글과 구별됨)
- 3-gate 통과 근거: 전부 직접 만들고 재현 가능(양식/수식), 전부 다운로드용 .xlsx 있음, 세금·보험·법률
  판단 없음 (packing slip 은 세금 계산 없이 수량만 다룸).

### 환경 문제 발견: LibreOffice Calc 컴포넌트 자체가 없었음
- `soffice --headless --convert-to xlsx ...` 를 새 템플릿 5개는 물론 기존 템플릿(craft-inventory 등)에도
  돌렸더니 전부 `Error: source file could not be loaded` (exit 0, 파일은 무시됨).
- `strace -f` 로 확인: `/usr/lib/libreoffice/program/` 안에 Calc 필터 라이브러리(libsc*, libscfilt* 등)가
  아예 없음. `dpkg -l | grep libreoffice` 결과 `libreoffice-core`/`libreoffice-common` 만 설치돼 있고
  `libreoffice-calc` 패키지가 없었음 (Writer 도 마찬가지로 없음).
- `apt-get update && apt-get install -y libreoffice-calc` 로 설치 (Ubuntu noble 저장소, 정상 설치됨).
  설치 후 동일 명령으로 기존 템플릿(craft-inventory-tracker.xlsx)과 새 템플릿 5개 모두 정상 변환 확인.
- 참고: 이전 브리프(2026-09-25)의 "재계산 결과 일치 확인"은 이 세션에서 직접 재현하지 않았음 — 그때는
  Calc 가 설치돼 있었거나 다른 방식으로 확인했을 가능성. 이번에 다시 5개 새 템플릿으로 직접 확인함.

### 검증 방법 및 결과 (전부 `soffice --headless --convert-to xlsx --outdir /tmp/recalc templates/<file>.xlsx`
로 강제 재계산 후 `openpyxl.load_workbook(..., data_only=True)` 로 계산된 값을 읽음)

**1. tidy-tabs-dependent-dropdown-list.xlsx** — Order form 탭 D열에 테스트용
`=COUNTA(INDIRECT(A2))` 를 Category 4개(Candles/Soap/Jewelry/Cards) 각각에 넣어, INDIRECT 가
Lists 탭의 이름정의 범위를 정확히 찾는지 확인 (드롭다운 자체 클릭은 UI 조작이라 이 세션에서 재현 불가,
그 아래 수식으로 대리 검증). 재계산 결과: Candles→3, Soap→2, Jewelry→4, Cards→2. 전부 Lists 탭
실제 항목 개수와 일치.

**2. tidy-tabs-client-contact-list.xlsx** — Next follow-up `=C2+D2`, Status
`=IF(E2="","",IF(TODAY()>=E2,"Follow up now","OK"))`. 실행일 2026-09-28 기준 재계산 결과:
Dana Ruiz(마지막 연락 09/01 +21일=09/22)→"Follow up now", Emma Lee(09/20+14=10/04)→"OK",
Sam Patel(08/15+30=09/14)→"Follow up now", Jamie Chen(09/25+7=10/02)→"OK". 기대값과 일치.
TODAY() 기반이라 이 결과는 2026-09-28 에 연 경우에만 유효하고, 다른 날 열면 값이 바뀜(의도된 동작).

**3. tidy-tabs-packing-slip-template.xlsx** — Total packed `=SUM(D11:D14)`. 재계산 결과 9
(2+3+1+3, 기대값과 일치). 인쇄 설정도 확인: `soffice --headless --convert-to pdf` 로 내보낸 뒤
`pdfinfo` 로 페이지 확인 — 워크북 전체(2번째 탭 "How to use" 포함)는 2쪽이지만, "How to use" 탭을
제외하고 Packing slip 탭만 남긴 사본을 변환하면 1쪽, 612x792pt(Letter), portrait. 즉 실제 인쇄 대상인
Packing slip 탭 자체는 US Letter 1쪽 그대로 나옴.

**4. tidy-tabs-running-balance-cash-log.xlsx** — Balance `=E{전행}+C{행}-D{행}`, 시작 잔액 250.00 은
값으로 직접 입력. 재계산 결과: 250 → 430(+180) → 384.5(-45.5) → 480.7(+96.2) → 445.7(-35). 손계산과
전부 일치.

**5. tidy-tabs-highlight-duplicates-sample.xlsx** — C열 `=IF(COUNTIF($B$2:$B$16,B2)>1,"Duplicate","")`,
조건부 서식도 같은 COUNTIF 수식 사용. 재계산 결과: priya.nair@example.com(3회), sam.patel@example.com(3회),
dana.ruiz@example.com(2회), ana.gomez@example.com(2회, 연속 두 줄 — 실수로 같은 주문 두 번 입력한
사례)이 전부 "Duplicate" 로 표시되고, 나머지 7개 고유 이메일은 빈 값. 기대한 대로.

Excel·Google Sheets 에서는 이번에도 직접 확인하지 않음 — team-topics 문서와 앞으로 나올 글의
`tested_in`/`Works in` 은 "LibreOffice Calc 24.2.7 (Linux) 만" 으로 정확히 표기할 것. Google Sheets 단계는
Google 도움말 링크로 근거를 댄다.

## 2026-09-26 — 이미지 기준 조사 (researcher, 공식 문서만)
- Google Images: 파일명은 짧고 설명적으로, alt 가 가장 중요, 이미지 주변에 관련 텍스트,
  alt 키워드 남용은 스팸 신호. https://developers.google.com/search/docs/appearance/google-images
- Discover: 가로 1200px 이상·16:9, max-image-preview:large, 로고·텍스트 과다·낚시 이미지 금지.
  https://developers.google.com/search/docs/appearance/google-discover
- AI 이미지 자체는 스팸 아님. 가치 없이 대량 생성하면 scaled content abuse. 제작 방식 표시 권장.
  https://developers.google.com/search/docs/fundamentals/using-gen-ai-content
- 구글이 "스크린샷이 사진보다 낫다"고 말한 공식 문구는 없음. 대표성·관련성 기준상 실제 결과
  이미지가 더 맞는다는 건 우리 판단.
- Threads: 이미지 글엔 링크 카드(link_attachment) 불가 → 클릭 목적이면 텍스트 + 링크 카드 유지.
  https://developers.facebook.com/docs/threads/posts
- 적용: hero(실제 결과, 1600x900) 먼저 → 단계 스크린샷 → AI 사진 1장(아래쪽, "AI-generated photo"
  캡션), 설명적인 파일명, 이미지별 alt, Threads 는 텍스트 + link_attachment.

## 2026-09-29 — 신규 주제 5개 (researcher) + 훅 + 템플릿 검증 (LibreOffice Calc 24.2.7, Linux)
- tasks/search-stats.md: 데이터 없음. 그래서 3-gate(직접 만들어 확인 가능 / 다운로드 양식 / 세금·법률 판단 없음)만 적용.
- 중복 확인: posts/ready, posts/published, topics-seed.md, team-topics-2026-09-29.md 의 제목과 겹치지 않음.
  (timesheet 글은 초과근무 계산이라 "시각 -> 소수 시간 변환" 은 별개. invoice/overdue 글과 "미수금 연령 보고서" 도 별개.)
- 검색 1페이지 예상 경쟁(추정, 미검증): Microsoft/Google 공식 도움말, ExcelJet/Ablebits 류. 롱테일(소상공인 예시, 다운로드 파일)로 비튼다.
- 순위(좋은 순): 1 receivables aging, 2 time to decimal hours, 3 markup vs margin, 4 business days ship-by, 5 VLOOKUP price list.

### 1. receivables-aging-report-excel-google-sheets
- 제목: How to Build an Accounts Receivable Aging Report in Excel and Google Sheets
- 첫 문단: Sort unpaid invoices by how late they are with `=MAX(0,AsOfDate-DueDate)`, label each one Current, 1-30, 31-60, 61-90, or 90+, and total each bucket with SUMIFS.
- 검색어 예: "accounts receivable aging report excel template", "aging report google sheets"
- 템플릿: templates/tidy-tabs-receivables-aging-report.xlsx (탭: Invoices, Aging, How to use)
- 검증(재계산 후 openpyxl data_only): 기준일 09/30/2026 고정. Days past due 47/21/21/0/103/62/0/0 (1001 은 Paid 라 공백). Bucket: 31-60, 1-30, 1-30, Current, 90+, 61-90, Current, Current. Aging 합계: Current $2,065.00 (3건), 1-30 $680.00 (2), 31-60 $1,200.00 (1), 61-90 $220.00 (1), 90+ $450.00 (1), 총 $4,615.00 (8). 손계산과 일치.

### 2. time-to-decimal-hours-excel-google-sheets
- 제목: How to Convert Clock Times to Decimal Hours in Excel and Google Sheets
- 첫 문단: Multiply the time difference by 24: `=ROUND(MOD(C2-B2,1)*24,2)` turns 8:30 AM to 5:00 PM into 8.5 hours, and MOD keeps overnight shifts positive.
- 검색어 예: "convert time to decimal hours excel", "overnight shift hours formula"
- 템플릿: templates/tidy-tabs-time-to-decimal-hours.xlsx (탭: Shifts, How to use)
- 검증: Time worked=MOD(C-B,1)-break/1440. 5행: 8:00 / 8:15 / 7:55 / 8:00(야간 10:00 PM-6:30 AM) / 4:15. Decimal: 8, 8.25, 7.92, 8, 4.25 (합 36.42). Nearest 15 min(MROUND): 8, 8.25, 8, 8, 4.25. Pay(시급 $22.50, ROUND 2): 180.00, 185.63, 178.20, 180.00, 95.63, 합 $819.46. 손계산 일치. 반올림 정책은 독자 몫이라고 글에 명시.

### 3. markup-vs-margin-calculator-excel-google-sheets
- 제목: How to Calculate Markup and Profit Margin in Excel and Google Sheets
- 첫 문단: Price from markup with `=B2*(1+C2)` and margin with `=(D2-B2)/D2`; to hit a target margin use `=B2/(1-G2)`, not `=B2*(1+G2)`.
- 검색어 예: "markup vs margin excel formula", "calculate selling price from margin"
- 템플릿: templates/tidy-tabs-markup-vs-margin-calculator.xlsx (탭: Pricing, How to use)
- 검증: 4행. 마크업 100/150/200/300% -> 가격 $8.20/$4.50/$10.20/$3.80, 이익 $4.10/$2.70/$6.80/$2.85, 마진 50.0%/60.0%/66.7%/75.0%. 목표마진 50/60/70/75% -> 가격 $8.20/$4.50/$11.33/$3.80; Margin check 0.5/0.6/0.6999/0.75 (11.33 반올림 때문에 69.99%, 표시는 70.0%). 손계산 일치.

### 4. business-days-ship-by-date-excel-google-sheets
- 제목: How to Add Business Days to a Date in Excel and Google Sheets
- 첫 문단: Use `=WORKDAY(C2,D2,Holidays!$A$2:$A$20)` to add business days to an order date; it skips weekends and every date in your holiday list.
- 검색어 예: "add business days to date excel", "workday formula google sheets exclude holidays"
- 템플릿: templates/tidy-tabs-business-days-ship-by-date.xlsx (탭: Orders, Holidays, How to use)
- 검증: 4행. 09/28+5 -> 10/05/2026 (7일), 10/06+10 -> 10/21/2026 (10/12 휴일 건너뜀, 15일), 11/09+3 -> 11/13/2026 (11/11 건너뜀, 4일), 11/20+5 -> 11/30/2026 (11/26 건너뜀, 10일). NETWORKDAYS(...)-1 검산이 각각 5/10/3/5 로 입력값과 일치. 손으로 달력 세어 확인.

### 5. vlookup-price-list-excel-google-sheets
- 제목: How to Look Up a Price With VLOOKUP in Excel and Google Sheets
- 첫 문단: Type `=VLOOKUP(A2,'Price list'!$A$2:$C$500,3,FALSE)` to pull the price for a SKU; FALSE forces an exact match, and wrapping it in IFERROR shows a clear message when the SKU is missing.
- 검색어 예: "vlookup exact match price list", "vlookup returns n/a"
- 템플릿: templates/tidy-tabs-vlookup-price-list.xlsx (탭: Order lines, Price list, How to use)
- 검증: 6행. Unit price 18/8/24/(못 찾음)/6/18, Line total 36/24/24/공백/30/18, Order total $132.00. "SOP-OAT " (끝 공백) 행은 "SKU not found" -> 트러블슈팅 재료. INDEX/MATCH 열이 VLOOKUP 열과 같은 값. XLOOKUP 은 LO 24.2 에서 쓰지 않음.

### 검증 방법과 한계
- 방법: `soffice --headless --convert-to xlsx --outdir <scratch> templates/<file>.xlsx` 후 openpyxl data_only=True 로 값 읽기.
- LibreOffice Calc 24.2.7 (Linux) 에서만 확인. Excel, Google Sheets 에서는 열어보지 않음 (글의 tested_in 은 LibreOffice 만).
- 환경: 이 세션에서도 libreoffice-calc 패키지가 없어서 변환이 실패했고 `apt-get install -y libreoffice-calc` 후 성공.
- build-templates.py 를 다시 돌리면 기존 템플릿 .xlsx 도 다시 저장돼 바이너리가 바뀜(내용 동일). 이번엔 git checkout 으로 기존 파일은 되돌림.

## 2026-10-02 — 신규 주제 5개 템플릿 제작/검증 (LibreOffice Calc 24.2.7, Linux)
- 주제 출처: tasks/team-topics-2026-10-02.md. 훅(제목, 첫 문단): tasks/hooks-2026-10-02.md.
- 빌드: scripts/build-templates.py 에 함수 5개 추가 (break_even_calculator, subscription_renewal_tracker, budget_vs_actual, farmers_market_sales_log, clean_customer_list) 후 main 에 등록.
  기존 템플릿 바이너리가 바뀌지 않도록 새 함수 5개만 import 해서 실행함 (전체 재실행 안 함).
- 검증 방법: `soffice --headless --convert-to xlsx --outdir <scratch> templates/<file>.xlsx` 후 openpyxl data_only=True 로 읽어서, 별도로 손/파이썬으로 계산한 기대값과 셀 단위로 비교.
  LibreOffice Calc 24.2.7 (Linux) 에서만 확인. Excel, Google Sheets 에서는 열어보지 않음 (글의 tested_in 은 LibreOffice 만).

### 1. templates/tidy-tabs-break-even-calculator.xlsx (탭: Break-even, How to use)
- 수식 (2행): Profit per unit `=B2-C2`, Break-even units `=IF(E2<=0,"No break-even",ROUNDUP(D2/E2,0))`, Break-even revenue `=IF(ISNUMBER(F2),F2*B2,"")`.
- 검증: 4행. 고정비/가격/변동비 -> units / revenue: 캔들 $120/$18.00/$6.50 -> 11 / $198.00 (120/11.5=10.43), 비누 $75/$8.00/$3.25 -> 16 / $128.00 (15.79), 귀걸이 $150/$24.00/$9.60 -> 11 / $264.00 (10.42), 카드 $60/$6.00/$1.50 -> 14 / $84.00 (13.33). 손계산과 일치.
- 가드 확인: 가격을 변동비보다 낮게(캔들 $6.00 < $6.50) 바꾼 사본을 재계산하면 units "No break-even", revenue 공백.

### 2. templates/tidy-tabs-subscription-renewal-tracker.xlsx (탭: Subscriptions, How to use)
- 수식 (2행): Next renewal `=IF(B2="","",EDATE(B2,C2))`, Days left `=IF(E2="","",E2-$K$1)`, Status `=IF(F2="","",IF(F2<0,"Past due",IF(F2<=30,"Renews soon","OK")))`, Monthly `=IF(D2="","",D2/C2)`, Annual `=IF(D2="","",D2*12/C2)`. K1 = 기준일 고정 10/02/2026 (TODAY() 쓰려면 K1 에 `=TODAY()` 입력, 노트에 적어둠). K2 `=SUM(H2:H200)`, K3 `=SUM(I2:I200)`.
- 검증 (기준일 10/02/2026): 도메인 03/15/2026+12개월 -> 03/15/2027, 164일, OK; 이메일 플랜 09/15+1 -> 10/15/2026, 13일, Renews soon; 쇼핑몰 플랜 09/28+1 -> 10/28/2026, 26일, Renews soon; 회계 소프트웨어 11/05/2025+12 -> 11/05/2026, 34일, OK; 사진 저장소 08/31/2026+6 -> 02/28/2027 (월말 보정), 149일, OK. 월 비용 $1.50/$20.00/$10.00/$15.00/$5.00 (합 $51.50), 연 비용 $18.00/$240.00/$120.00/$180.00/$60.00 (합 $618.00). 손계산 일치.
- EDATE 월말: `=EDATE(DATE(2026,1,31),1)` = 02/28/2026 (별도 임시 파일로 확인). 08/31+6개월 -> 02/28/2027 은 샘플 행으로 확인.
- 처음 내 기대값이 틀렸던 것: 회계 소프트웨어 34일은 30일 초과라 "OK" 가 맞음 (기대값을 "Renews soon" 으로 잘못 적었다가 수정, 수식 버그 아님).

### 3. templates/tidy-tabs-budget-vs-actual.xlsx (탭: Budget vs actual, How to use)
- 수식 (2행): Variance $ `=C2-B2`, Variance % `=IF(B2=0,"",(C2-B2)/B2)`, Status `=IF(C2>B2,"Over budget","On or under")`. 합계는 표 옆 H1:I4: `=SUM(B2:B200)`, `=SUM(C2:C200)`, `=I2-I1`, `=IF(I1=0,"",(I2-I1)/I1)`. 조건부 서식 `AND($C2<>"",$C2>$B2)` 로 초과 행 빨강.
- 검증 (예산 -> 실제 = 차이, %): Booth $450.00->$450.00 = $0.00, 0.0%; Materials $600.00->$683.40 = $83.40, 13.9%; Shipping $220.00->$241.75 = $21.75, 9.9%; Packaging $120.00->$131.20 = $11.20, 9.3%; Software $85.00->$85.00 = $0.00; Advertising $150.00->$92.50 = -$57.50, -38.3%; Payment fees $90.00->$87.35 = -$2.65, -2.9%; Training $0.00->$45.00 = $45.00, % 공백(0 예산 케이스, 상태 Over budget). 합계 예산 $1,715.00, 실제 $1,816.20, 차이 $101.20, 5.9%. 손계산 일치.
- 합계는 "행"이 아니라 표 옆 블록 (기존 인벤토리 템플릿 관례, 행 추가 시 안 밀림).

### 4. templates/tidy-tabs-farmers-market-sales-log.xlsx (탭: Sales, Daily totals, How to use)
- 수식: Line total `=C2*D2`. Daily totals 2행: Revenue `=SUMPRODUCT((Sales!$A$2:$A$500=A2)*Sales!$C$2:$C$500*Sales!$D$2:$D$500)`, 검산 `=SUMIFS(Sales!$E$2:$E$500,Sales!$A$2:$A$500,A2)`, Items sold `=SUMIFS(Sales!$C$2:$C$500,Sales!$A$2:$A$500,A2)`. 전체: G1 `=SUMPRODUCT(Sales!C2:C500,Sales!D2:D500)`, G2 `=SUM(Sales!E2:E500)`.
- 검증: 09/12/2026 $418.00 (38개), 09/19/2026 $482.00 (37개), 09/26/2026 $410.00 (37개); SUMPRODUCT 와 SUMIFS 열 동일. 전체 $1,310.00 (G1, G2 모두). 손계산 (예: 9/12 = 9x18+14x8+3x24+12x6) 일치.
- 한계: Qty/Price 셀에 텍스트가 있으면 SUMPRODUCT 곱셈식은 #VALUE! 가 날 수 있음 (노트에 적음). 이 상황은 재현 테스트 안 함.

### 5. templates/tidy-tabs-clean-customer-list.xlsx (탭: Clean names, How to use)
- 수식 (2행): B `=PROPER(TRIM(CLEAN(A2)))`, C `=PROPER(TRIM(CLEAN(SUBSTITUTE(SUBSTITUTE(A2,UNICHAR(160)," "),CHAR(10)," "))))` (파일에는 `_xlfn.UNICHAR` 로 저장), D `=LEN(A2)`, E `=LEN(C2)`.
- 버그 발견/수정: 처음엔 brief 대로 `CHAR(160)` 을 썼는데, LibreOffice 에서 CHAR(160) 은 시스템 로케일에 따라 달라짐. POSIX 로케일에서는 nbsp 로 동작했지만 UTF-8 로케일(LC_ALL=C.UTF-8, python subprocess 기본)에서는 `CODE(CHAR(160))` = 239 가 나와서 nbsp 가 안 지워짐 (Sam\xa0Patel 유지). `UNICHAR(160)` 으로 바꾸니 두 로케일 모두 정상. 글에서 CHAR(160) 을 쓰려면 이 LibreOffice 한계를 적거나 UNICHAR(160) 을 쓸 것. Excel/Sheets 의 CHAR(160) 동작은 확인 안 함.
- 검증 (열 C 기준, 원본 -> 결과): "  dana   ruiz " -> Dana Ruiz; "EMMA LEE" -> Emma Lee; "sam"+nbsp+"patel" -> Sam Patel; "jamie"+줄바꿈+"chen" -> Jamie Chen; "  PRIYA   NAIR  " -> Priya Nair; "ana gomez" -> Ana Gomez; "ronald McDONALD" -> Ronald Mcdonald (엣지); "sean o'neil" -> Sean O'Neil; "TOM  REYES"+nbsp -> Tom Reyes; "   lee wu" -> Lee Wu.
  길이 열 D(before)/E(after), 행 2~11: before 14,8,9,10,16,9,15,11,11,9 / after 9,8,9,10,10,9,15,11,9,6. 전부 기대값과 일치 (TOM 행: 원본 11자 -> "Tom Reyes" 9자).
- 열 B (기본 공식)의 한계도 확인: nbsp 행은 "Sam"+nbsp+"Patel" 그대로, 줄바꿈 행은 "Jamiechen" (CLEAN 이 줄바꿈을 공백 없이 삭제, PROPER 가 뒤 단어를 소문자로), TOM 행은 끝 nbsp 남음. 트러블슈팅 재료.
- McDonald -> "Mcdonald", o'neil -> "O'Neil" 둘 다 LibreOffice 에서 기대대로 나옴.

## 2026-10-06 — 신규 주제 5개 템플릿 제작/검증 (LibreOffice Calc 24.2.7, Linux)
- 주제: tasks/team-topics-2026-10-06.md. 훅: tasks/hooks-2026-10-06.md.
- 빌드: scripts/build-templates.py 에 함수 5개 추가, 새 함수만 실행 (기존 바이너리 변경 없음).
- 검증: `soffice --headless --convert-to xlsx` 후 openpyxl data_only=True. LibreOffice 에서만 확인. Excel, Google Sheets 에서는 열어보지 않음.

### 1. templates/tidy-tabs-loan-payment-calculator.xlsx (탭: Loans, How to use)
- 수식: Monthly `=PMT(C2/12,D2*12,-B2)`, Total paid `=E2*D2*12`, Total interest `=F2-B2`.
- 검증: $15,000.00 / 6.50% / 5년 -> $293.49 (293.4922), 총 $17,609.53, 이자 $2,609.53; $8,000 / 7.25% / 3년 -> $247.93, $8,925.56, $925.56; $22,000 / 5.90% / 7년 -> $320.33, $26,908.10, $4,908.10; $5,000 / 9.00% / 2년 -> $228.42, $5,482.17, $482.17. 첫 행 손계산 일치.

### 2. templates/tidy-tabs-percent-change.xlsx (탭: Sales, How to use)
- 수식: `=C2-B2`, `=IF(B2=0,"",(C2-B2)/B2)`.
- 검증: 2400->2760 = $360.00, 15.0%; 1800->1692 = -$108.00, -6.0%; 950->1140 = $190.00, 20.0%; 0->300 = $300.00, % 공백; 640->640 = $0.00, 0.0%.

### 3. templates/tidy-tabs-rank-top-customers.xlsx (탭: Customers, How to use)
- 수식: Rank `=RANK(B2,$B$2:$B$11,0)`, Unique `=C2+COUNTIF($B$2:B2,B2)-1`, Top 3 `INDEX/MATCH` on unique rank.
- 검증: 10행 (표 순서). Rank 열 = 1,2,2,5,6,6,8,4,10,9; Unique 열 = 1,2,3,5,6,7,8,4,10,9. 동점 $3,920.00 x2 -> rank 2,2 / unique 2,3 (rank 3 은 없음); $2,150.00 x2 -> rank 6,6 / unique 6,7. Lakeview $3,100.00 = rank 4. Top 3 = Harbor Coffee $4,850.00, Maple Street Bakery $3,920.00, Oak & Ember Candles $3,920.00. 손으로 순위 세어 일치.

### 4. templates/tidy-tabs-invoice-number-generator.xlsx (탭: Invoices, How to use)
- 수식: C `="INV-"&TEXT(B2,"yyyy")&"-"&TEXT(ROWS($A$2:A2),"0000")`, D `="INV-"&YEAR(B2)&"-"&TEXT(ROWS($A$2:A2),"0000")`.
- 검증: 6행, 날짜 09/28/2026 ~ 10/12/2026 -> INV-2026-0001 ~ INV-2026-0006, C 와 D 열 동일. TEXT "yyyy" 의 로케일 의존성은 재현 안 함 (노트에 YEAR 대안만 안내).

### 5. templates/tidy-tabs-combine-address-columns.xlsx (탭: Addresses, How to use)
- 수식: F `=_xlfn.TEXTJOIN(", ",TRUE,A2:D2)` (Excel 에서는 TEXTJOIN), G `=F2&" "&E2` 형태 (TEXTJOIN(...)&" "&E2), H `=A2&", "&B2&", "&C2&", "&D2`.
- 검증: 5행. 단위(Unit) 있는 행 "118 Elm Street, Apt 4B, Boston, MA", 없는 행 "2450 Harbor Road, Portland, ME", ZIP 포함 "..., ME 04101" (ZIP 은 텍스트라 02108 유지). H 열은 단위가 비면 "2450 Harbor Road, , Portland, ME" (이중 쉼표). 
- 한계: Excel 2016 이하에서 TEXTJOIN 미지원은 Microsoft 문서 기준 (직접 확인 안 함).

## 2026-10-07 — 신규 주제 5개 템플릿 제작/검증 (LibreOffice Calc 24.2.7, Linux)
- 주제: tasks/team-topics-2026-10-07.md. 훅: tasks/hooks-2026-10-07.md.
- 빌드: scripts/build-templates.py 에 함수 5개 추가 (sale_price_discount, shipping_weight_tier_lookup, reorder_point_calculator, freelance_hourly_rate, convert_text_to_numbers) 후 main 에 등록. 새 함수 5개만 import 해서 실행 (기존 바이너리 변경 없음). 참고: 2026-10-06 함수 5개(loan_payment 등)는 main 에 아직 등록 안 돼 있음 (이번엔 손대지 않음).
- 검증: `soffice --headless --convert-to xlsx --outdir <scratch> templates/<file>.xlsx` 후 openpyxl data_only=True 로 읽고 손계산 기대값과 셀 단위 비교. **LibreOffice Calc 24.2.7 (Linux) 에서만 확인. Excel, Google Sheets 에서는 열어보지 않음.** FILTER/LET 미사용.
- 환경: python3 (/usr/local/bin) 에는 openpyxl 이 없고 /usr/bin/python3 에 있음.

### 1. templates/tidy-tabs-sale-price-discount-calculator.xlsx (탭: Sale prices, Reverse and stacked, How to use)
- 수식: Sale price `=B2*(1-C2)`, You save `=B2-D2`, Original price (역산) `=D2/(1-C2)`, 중복할인 `=B8*(1-B9)*(1-B10)`, 잘못된 합산 `=B8*(1-(B9+B10))`.
- 검증: 8행 판매가 $16.20/$15.30/$6.00/$7.65/$20.40/$4.50/$36.00/$28.80, 절약 $1.80/$2.70/$2.00/$0.85/$3.60/$1.50/$12.00/$3.20; 합계 원가 $162.50, 판매가 $134.85, 절약 $27.65. $48.00 25% 할인 = $36.00. 역산 $36.00 @25% -> $48.00, $20.40 @15% -> $24.00, $7.65 @10% -> $8.50. $100.00 에 20% 후 10% = $72.00 (30% 합산은 $70.00, 차이 $2.00, 실제 총 할인 28%). 손계산 일치.

### 2. templates/tidy-tabs-shipping-weight-tier-lookup.xlsx (탭: Orders, Unsorted table demo, How to use)
- 수식: `=IFERROR(VLOOKUP(B2,$G$2:$H$7,2,TRUE),"Below minimum")`, `=IFERROR(INDEX($H$2:$H$7,MATCH(B2,$G$2:$G$7,1)),"Below minimum")`, 오류 확인용 `=VLOOKUP(B2,$G$2:$H$7,2,TRUE)`. 요율표(From oz -> $): 1 -> 4.50, 4 -> 5.75, 8 -> 7.25, 16 -> 9.80, 32 -> 13.40, 64 -> 18.90 (가짜 요율).
- 검증: 12행 무게 2.5/3.99/4/4.01/7.9/8/12.5/16/20/40/64/0.5 -> $4.50/$4.50/$5.75/$5.75/$5.75/$7.25/$7.25/$9.80/$9.80/$13.40/$18.90/"Below minimum". VLOOKUP 과 INDEX/MATCH 열 동일. 경계: 4 oz = $5.75 (4 oz 구간 시작), 3.99 = $4.50, 4.01 = $5.75. 0.5 oz 는 IFERROR 없이 #N/A.
- 정렬 안 된 표 데모 (1, 8, 4, 16, 32, 64 순서): 5 oz -> $5.75 (우연히 맞음), 10 oz -> $5.75 (정답 $7.25, 틀림, 오류 표시 없음). 이 동작은 LibreOffice 에서만 확인. Excel/Sheets 가 정렬 안 된 표에서 같은 값을 줄지는 확인 안 함 (노트에도 그렇게 적음).

### 3. templates/tidy-tabs-reorder-point-calculator.xlsx (탭: Reorder, How to use)
- 수식: Daily `=B2/$K$1` (K1 = 30), Reorder point `=ROUNDUP(C2*D2+E2,0)`, Flag `=IF(G2<=F2,"REORDER","OK")`. 조건부 서식으로 REORDER 행 빨강.
- 검증: 6행. 90개/7일/안전 6 -> 일 3.00, 재주문점 27, 재고 40 OK. 60/10/8 -> 2.00, 28, 재고 28 REORDER (경계, 같을 때 REORDER). 33/14/4 -> 1.10, 20 (19.4 올림), 재고 21 OK. 75/5/10 -> 2.50, 23 (22.5 올림), 재고 18 REORDER. 120/12/15 -> 4.00, 63, 재고 80 OK. 45/21/5 -> 1.50, 37 (36.5 올림), 재고 30 REORDER. 손계산 일치.
- 주제 파일은 craft inventory tracker 글 링크를 요구함 (본문에서 연결, 템플릿엔 링크 없음).

### 4. templates/tidy-tabs-freelance-hourly-rate-calculator.xlsx (탭: Hourly rate, How to use)
- 수식: Weeks worked `=52-B4`, Hours worked `=B7*B5`, Billable hours `=B8*B6`, Money needed `=B2+B3`, Rate `=B10/B9`, 올림 `=ROUNDUP(B11,0)`. 입력은 B2:B6, 시나리오 3열(C:E)은 같은 수식.
- 검증: 기본 (60,000 + 6,000, 4주 휴가, 주 30시간, 65%) -> 48주, 1,440시간, 936 청구시간, $66,000.00, $70.51 (70.5128), 올림 $71. What-if 1 (목표 $75,000.00): $81,000.00 / 936 = $86.54. What-if 2 (휴가 6주): 46주, 1,380시간, 897, $73.58. What-if 3 (청구 75%): 1,080, $61.11. 손계산 일치. 세금/자영업 비용 미포함, 노트에 회계사 안내 문구 넣음.

### 5. templates/tidy-tabs-convert-text-to-numbers.xlsx (탭: Pasted amounts, How to use)
- 수식: ISNUMBER `=ISNUMBER(A2)`, Basic clean `=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"$",""),",",""))`, 괄호 처리 `=IF(LEFT(TRIM(A2),1)="(",-1,1)*VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),"$",""),",",""),"(",""),")",""))`. A열은 텍스트 서식(@)로 문자열 저장.
- 검증: 10행 전부 ISNUMBER FALSE (텍스트). 텍스트 열 SUM = 0. Basic clean 합 = 괄호 처리 열 합 = $5,095.60 (손계산 1250+18-45.5+89.95+2400+12.5+7.25+1075.4+310-22). 정리 후 ISNUMBER 전부 TRUE.
- 발견: LibreOffice 에서는 Basic clean(괄호 처리 없음)도 VALUE("(45.50)") = -45.5, VALUE("(22.00)") = -22 로 나옴. Excel/Sheets 의 VALUE 가 괄호를 어떻게 읽는지는 확인 안 함 -> 글에서는 괄호 처리 열(D)을 쓰고, LibreOffice 동작에 기대지 말라고 쓸 것 (노트에도 적음).
- 로케일: US 로케일에서만 확인. 다른 지역 설정의 $ , . 해석은 미확인.

## 2026-10-08 #2 — templates/tidy-tabs-sum-expenses-by-category.xlsx (탭: Expenses, How to use)
- 빌드: build-templates.py 의 sum_expenses_by_category() 만 import 해서 실행. LibreOffice Calc 24.2.7 (Linux) 에서만 확인, Excel/Google Sheets 는 열어보지 않음.
- 수식: G2:G8 `=SUMIF($C$2:$C$40,F2,$D$2:$D$40)`, H `=G2/$G$9`, G9 `=SUM(G2:G8)`, 확인 G11 `=SUM($D$2:$D$40)`, G12 `=ROUND(G11-G9,2)`, G13 `=IF(G12=0,"OK","CHECK CATEGORIES")`.
- 검증: 로그 22행. Supplies $317.30 (22.2%), Software $90.99 (6.4%), Shipping $142.25 (9.9%), Travel $246.00 (17.2%), Meals $74.35 (5.2%), Utilities $240.74 (16.8%), Advertising $320.50 (22.4%). Total $1,432.13 = 로그 합계, 차이 $0.00, Check OK. 퍼센트 합은 100.0% (표시값 반올림 합은 100.1%). 손계산 일치.
- 함정 재현 (임시 복사본): C2 를 "Supplies " (끝 공백) 으로 바꾸면 LibreOffice 에서 Supplies 합계 $232.80 (=$317.30 - $84.50, 해당 행 누락), Total $1,347.63, 로그 합계 $1,432.13, 차이 $84.50, Check = CHECK CATEGORIES. 오류 표시 없음. Excel/Sheets 동작은 확인 안 함.

## 2026-10-08 — 주제 5: year-to-date sales (LibreOffice Calc 24.2.7, Linux)

### templates/tidy-tabs-year-to-date-sales-total.xlsx (탭: Sales, Text date demo, How to use)
- 수식: YTD `=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),1,1),A2:A500,"<="&F1)`, MTD 은 시작을 `DATE(YEAR(F1),MONTH(F1),1)`, 월말까지는 `EOMONTH(F1,0)`. F1 = 보고 날짜 (고정 08/31/2026, TODAY 미사용). 이 파일은 FILTER/LET 미사용.
- 검증 (LibreOffice 재계산, 손계산 일치): 14행. F1=08/31/2026 -> YTD $9,596.75, MTD $1,655.50, 월 전체 $1,655.50 (08/31 행 $275.00 포함, 09/12 $1,190.00 와 10/03 $330.00 제외). F1=07/31/2026 -> YTD $7,941.25, MTD $1,465.00. F1=08/15/2026 -> YTD $8,461.75, MTD $520.50, 월 전체 $1,655.50. F1=12/31/2026 -> YTD $11,116.75, MTD $0.00.
- 텍스트 날짜 재현: Text date demo 탭 (08/18/2026 $860.00 을 텍스트로 저장) -> YTD $8,736.75 (정상 $9,596.75 보다 $860.00 적음), MTD $795.50 (정상 $1,655.50). 오류 표시 없이 조용히 빠짐. COUNT(A) = 13 vs COUNT(C) = 14 로 확인 가능. 이 동작은 LibreOffice 에서만 확인, Excel/Sheets 는 확인 안 함.
- LibreOffice Calc 24.2.7 (Linux) 에서만 확인. Excel, Google Sheets 에서는 열어보지 않음.

## 2026-10-08 — Topic 1: sales tax calculator (LibreOffice Calc 24.2.7, Linux)
### templates/tidy-tabs-sales-tax-calculator.xlsx (탭: Sales tax, Tax-inclusive totals, How to use)
- 수식: 세금 `=ROUND(B2*$F$1,2)` (F1 = 6.25%, 가짜 요율), 합계 `=B2+C2`, 합계행 SUM. 두 번째 탭: `=ROUND(B2-B2/(1+$F$1),2)` (세금 역산), `=B2-C2`; 요율은 `='Sales tax'!F1`.
- 검증: 6행 가격 $18.00/$4.50/$24.00/$16.00/$12.99/$32.50 -> 세금 $1.13/$0.28/$1.50/$1.00/$0.81/$2.03, 합계 $19.13/$4.78/$25.50/$17.00/$13.80/$34.53. 합계행 $107.99 / $6.75 / $114.74. 역산 탭: 같은 총액 -> 원가 $18.00/$4.50/$24.00/$16.00/$12.99/$32.50 (왕복 일치), $50.00 -> 세금 $2.94, 원가 $47.06. 손계산 일치 (18 x 0.0625 = 1.125 -> 1.13, 즉 .5 에서 올림).
- **LibreOffice Calc 24.2.7 (Linux) 에서만 확인. Excel, Google Sheets 에서는 열어보지 않음.** 세금 자문 아님, 면세 품목 처리는 템플릿에 없음.
- 링크: Google 도움말은 curl -I (HEAD) 로는 404, GET 으로는 200 (ROUND 3093440 제목 확인). Microsoft ROUND 링크 HEAD 200.

## 2026-10-08 — 주제 3: round prices nearest nickel / .99
### templates/tidy-tabs-round-prices-nearest-nickel-99.xlsx (탭: Prices, How to use)
- 수식: C `=ROUND(B2,2)`, D `=MROUND(B2,0.05)`, E `=CEILING(B2,0.05)`, F `=ROUNDUP(B2,0)-0.01`. B 는 반올림 전 가격(4자리).
- 검증 (LibreOffice Calc 24.2.7 Linux 에서만; Excel, Google Sheets 에서는 열어보지 않음): 8행 13.4167 -> 13.42/13.40/13.45/13.99; 8.2333 -> 8.23/8.25/8.25/8.99; 24.7083 -> 24.71/24.70/24.75/24.99; 17.9625 -> 17.96/17.95/18.00/17.99; 6.13 -> 6.13/6.15/6.15/6.99; 31.4467 -> 31.45/31.45/31.45/31.99; 12.9833 -> 12.98/13.00/13.00/12.99; 22.0000 -> 22.00/22.00/22.00/21.99 (정수 가격은 .99 수식이 1센트 내림). 손계산 일치.
- 부호: 임시 파일에서 LibreOffice 는 MROUND(13.4167,-0.05) = 13.4 를 반환 (Excel 동작은 확인 안 함). 글에서 음수 동작은 주장하지 않음. 부동소수점 표시 문제(12.9899999)는 재현 안 함, 글에 쓰지 않음.

## 2026-10-08 — 주제 4: Remove non-breaking spaces (templates/tidy-tabs-remove-non-breaking-spaces.xlsx, 탭: Pasted names, How to use)
- 빌드: scripts/build-templates.py remove_non_breaking_spaces() (main 등록은 리드). 셀 A 에 Python chr(160), 끝 줄바꿈(\n), 끝 탭(\t) 직접 삽입.
- 수식: LEN `=LEN(A2)`, 코드 `=_xlfn.UNICODE(MID(A2,5,1))`, TRIM만 `=TRIM(A2)`, 정리 `=TRIM(CLEAN(SUBSTITUTE(A2,_xlfn.UNICHAR(160)," ")))`, COUNTIF/VLOOKUP(FALSE) 를 L2:M9 고객 목록에 대해 정리 전(A)/후(E).
- 검증 (LibreOffice Calc 24.2.7 Linux 만, Excel/Sheets 아님): 8행 LEN 전 9/8/9/10/11/10/8/10 -> 후 9/8/8/9/10/9/8/10. 5번째 문자 코드 160 (6행), 32 (Mark Diaz, Mike Stone). COUNTIF 전 합계 1 (Mike Stone 만), 후 8. VLOOKUP 전 7행 "not found", 후 $120.00/$45.00/$89.50/$210.00/$65.00/$150.00/$38.00/$72.25. LibreOffice TRIM 은 CHAR(160) 을 지우지 않음 (LEN 9 -> 9). 손계산 일치.
- 발견: LibreOffice 에서 CHAR(160)/CODE 는 프로필 로케일에 따라 다름. 기본 프로필 160 (정상), 새 프로필(-env:UserInstallation, render-previews 방식) 에서 CODE = 194 이고 SUBSTITUTE(CHAR(160)) 가 안 먹음. 그래서 템플릿은 UNICHAR/UNICODE 사용 (clean_customer_list 와 동일). 글은 CHAR(160) 을 답으로 쓰고 Troubleshooting 에 이 차이를 적음. Excel/Sheets 의 CHAR(160) 동작은 확인 안 함.
- Microsoft TRIM 페이지 본문(curl) 확인: "designed to trim the 7-bit ASCII space character (value 32)... By itself, the TRIM function does not remove this nonbreaking space character." Google 도움말 URL 은 HEAD 로는 404, GET 은 200 (TRIM 3094140, SUBSTITUTE 3094215, import 40608).

## 2026-10-08 — 신규 주제 5개 (researcher)
- tasks/search-stats.md 없음 -> **데이터 없음.** 숫자 기반 판단 불가, 3-gate (직접 제작/검증, 다운로드 가능, 세금·법률 판단 없음)만 적용. 제목/슬러그가 posts/ready, posts/published, tasks/topics-seed.md 와 겹치지 않는지 /tmp/titles.txt 로 대조함. 아래는 좋은 순. 제목 다듬기와 본문은 hook-writer 몫.
- 공통: 모든 검증은 .xlsx 를 LibreOffice (soffice --headless --convert-to xlsx) 로 재계산 후 openpyxl data_only=True 로 값 확인. Excel/Google Sheets 에서는 열어보지 않음 (글에서는 공식 도움말 링크로 근거).

### 1. ~~How to Convert Text Dollar Amounts From a CSV Into Real Numbers in Excel and Google Sheets~~ (삭제, 2026-10-09)
- 기존 글과 중복이라 제외: convert-text-to-numbers-dollar-signs 와 같은 주제. 계획 전문은 git 기록에 있음.

### 2. ~~How to Calculate Your Effective Hourly Rate on a Fixed-Price Project in Excel and Google Sheets~~ (삭제, 2026-10-09)
- 기존 글과 중복이라 제외: freelance-hourly-rate-calculator 와 겹침. 계획 전문은 git 기록에 있음.

### 3. How to Set Wholesale Price Tiers by Quantity With an Approximate-Match Lookup in Excel and Google Sheets
- slug: `quantity-price-tiers-lookup-excel-google-sheets`
- 검색어: "excel quantity discount tiers formula", "vlookup approximate match price breaks", "wholesale pricing tiers google sheets"
- 샘플/템플릿: tidy-tabs-quantity-price-tiers.xlsx (탭: Tiers, Orders, How to use). Tiers: 1-11 $12.00, 12-35 $10.50, 36-71 $9.25, 72+ $8.00 (임의 샘플 가격). Orders 8행: 수량 -> 단가 -> 합계.
- 수식: `=VLOOKUP(B2,Tiers!$A$2:$B$5,2,TRUE)`, `=B2*C2`, 대안 `=LOOKUP(B2,Tiers!$A$2:$A$5,Tiers!$B$2:$B$5)`, `=INDEX(Tiers!$B$2:$B$5,MATCH(B2,Tiers!$A$2:$A$5,1))`, 경계값 테스트용 열.
- 확인: 수량 1, 11, 12, 35, 36, 71, 72, 500 (경계) 에서 단가가 기대값. 세 방식이 동일한 값을 내는지. 오름차순 정렬이 깨지면 틀린 값이 나오는 것을 일부러 재현해서 "정렬 필수" 경고 근거로 쓴다. 0 수량 -> #N/A 처리 (IFERROR) 확인.
- 경쟁: 기존 `vlookup-price-list` 글은 정확히 일치(FALSE) 위주라 TRUE 근사 일치는 별개 주제. Microsoft VLOOKUP 도움말이 1페이지에 있음 -> "수량 할인/도매" 상황으로 롱테일.

### 4. How to Calculate Weighted Average Cost of Supplies With SUMPRODUCT in Excel and Google Sheets
- slug: `weighted-average-cost-sumproduct-excel-google-sheets`
- 검색어: "weighted average cost excel sumproduct", "average unit cost when price changes spreadsheet", "calculate average cost of inventory purchases google sheets"
- 샘플/템플릿: tidy-tabs-weighted-average-cost.xlsx (탭: Purchases, Summary, How to use). Purchases: 날짜 MM/DD/YYYY, 품목 (soy wax, wicks, jars), 수량, 단가, 합계. Summary: 품목별 가중평균 단가, 단순평균과 비교.
- 수식: `=C2*D2`, 품목별 `=SUMIFS(E:E,B:B,A2)/SUMIFS(C:C,B:B,A2)`, `=SUMPRODUCT((B2:B20=A2)*C2:C20*D2:D20)/SUMIFS(C2:C20,B2:B20,A2)`, `=AVERAGEIFS(D2:D20,B2:B20,A2)` (단순평균 비교).
- 확인: 손계산 (예: 10 lb x $4.00 + 40 lb x $3.50 = $180 / 50 = $3.60, 단순평균은 $3.75) 일치, SUMPRODUCT 와 SUMIFS 방식 값 동일. 회계기준(FIFO 등)·세금 처리 판단은 하지 않음 - "단순 평균 단가 계산"만.
- 경쟁: Exceljet/Corporate Finance 류 일반 "weighted average" 글. 공예/베이커리 재료 구매 상황으로 롱테일.

### 5. ~~How to Round Prices to .99 or the Nearest Quarter in Excel and Google Sheets~~ (삭제, 2026-10-09)
- 기존 글과 중복이라 제외: round-prices-nearest-nickel-99 와 같은 주제. 계획 전문은 git 기록에 있음.

## 2026-10-09 신규 주제 5개 (hook-writer + 템플릿 빌드)
- 리서처 10-08 목록 중 1, 2, 5번은 기존 글과 중복이라 삭제(위에 취소선). 3번(quantity-price-tiers)과 4번(weighted-average-cost)은 유지, 새로 3개 추가 (MAXIFS, COUNTIFS 월별, AVERAGEIFS 채널별). 검색 데이터 없음, 3-gate 기준만 적용.
- 템플릿은 scripts/build-templates.py 에 함수 5개 추가. 새 .xlsx 5개만 커밋 대상 (재빌드 시 기존 바이너리는 zip 타임스탬프만 바뀌어 git checkout 으로 되돌림).
- 검증 방법 공통: soffice --headless --convert-to xlsx 로 재계산 후 openpyxl data_only=True. **LibreOffice Calc 24.2 (Linux) 에서만 확인. Excel, Google Sheets 에서는 열어보지 않음.** 글의 tested_in 에 그대로 적을 것. MAXIFS 는 _xlfn.MAXIFS 로 저장했고 LibreOffice 가 정상 계산함.

### A. quantity-price-tiers-lookup-excel-google-sheets
- 제목: How to Look Up Wholesale Price Tiers by Quantity in Excel and Google Sheets
- 템플릿: tidy-tabs-quantity-price-tiers.xlsx (Tiers, Orders, Unsorted demo, How to use)
- 수식: `=VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE)`, `=LOOKUP(C2,Tiers!$A$2:$A$5,Tiers!$B$2:$B$5)`, `=INDEX(Tiers!$B$2:$B$5,MATCH(C2,Tiers!$A$2:$A$5,1))`, 합계 `=C2*D2`.
- 검증: 수량 1/11/12/35/36/71/72/500 -> $12.00/$12.00/$10.50/$10.50/$9.25/$9.25/$8.00/$8.00, 합계 $12.00/$132.00/$126.00/$367.50/$333.00/$656.75/$576.00/$4,000.00. 세 방식 8행 모두 일치. 수량 0: VLOOKUP #N/A, IFERROR 로 감싸면 "Below first tier". 손계산 일치.
- 정렬 깨진 표(36,1,72,12 순서): LibreOffice 는 수량 5 -> #N/A, 20 -> #N/A, 40 -> $12.00 (정답 $9.25), 80 -> $10.50 (정답 $8.00). 틀린 값이 나오는 건 확인했지만 정확한 결과는 구현에 따라 다를 수 있어 글에는 "틀리거나 #N/A" 로만 쓰고 Excel/Sheets 가 같다고 쓰지 않음.

### B. weighted-average-cost-sumproduct-excel-google-sheets
- 제목: How to Calculate Weighted Average Cost With SUMPRODUCT in Excel and Google Sheets
- 템플릿: tidy-tabs-weighted-average-cost.xlsx (Purchases, Summary, How to use)
- 수식: `=C2*D2`, `=SUMPRODUCT((Purchases!$B$2:$B$40=A2)*Purchases!$C$2:$C$40*Purchases!$D$2:$D$40)/B2`, `=C2/B2` (SUMIFS 합계 기반), `=AVERAGEIFS(Purchases!$D$2:$D$40,Purchases!$B$2:$B$40,A2)`.
- 검증: Soy wax 75 lb, $275.00, 가중 $3.6667, 단순 $3.7667. Cotton wicks 1,000, $99.00, 가중 $0.0990, 단순 $0.1033. Glass jars 168, $292.80, 가중 $1.7429, 단순 $1.8500. SUMPRODUCT 와 SUMIFS 값 동일. 손계산 일치. 재고평가법/세금 판단 아님을 노트에 적음.

### C. last-order-date-per-customer-maxifs-excel-google-sheets
- 제목: How to Find the Last Order Date for Each Customer With MAXIFS in Excel and Google Sheets
- 템플릿: tidy-tabs-last-order-date-maxifs.xlsx (Orders, Summary, How to use)
- 수식: `=MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2)` (파일에는 _xlfn.MAXIFS), 마지막 금액 `=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2,Orders!$A$2:$A$200,B2)`, 경과일 `=$H$2-B2` (H2 = 10/09/2026).
- 검증: Harbor Coffee 09/17/2026 $865.00 (22일), Maple Street Bakery 09/30/2026 $428.40 (9일), Oak & Ember Candles 09/09/2026 $510.75 (30일), Lakeview Florist 09/24/2026 $242.00 (15일). 손계산 일치. LibreOffice 24.2 는 MAXIFS 지원.
- 한계: 같은 날 주문이 2건이면 SUMIFS 가 합산함 (노트에 적음). 일치 없으면 MAXIFS 는 0 (LibreOffice 동작 기준으로만 서술, Excel/Sheets 는 확인 안 함). Excel 2016 이하에는 MAXIFS 없음 (Microsoft 도움말 링크로 근거, 직접 확인 안 함).

### D. count-orders-by-month-countifs-excel-google-sheets
- 제목: How to Count Orders by Month With COUNTIFS in Excel and Google Sheets
- 템플릿: tidy-tabs-count-orders-by-month.xlsx (Orders, Monthly, How to use)
- 수식: `=COUNTIFS(Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1))`, `=SUMIFS(Orders!$C$2:$C$200,Orders!$A$2:$A$200,">="&A2,Orders!$A$2:$A$200,"<"&EDATE(A2,1))`.
- 검증: 07/2026 2건 $365.50, 08/2026 3건 $549.25 (08/31 포함), 09/2026 3건 $693.75 (09/01 포함), 10/2026 2건 $395.50. 합계 10건 $2,004.00 = Orders 탭 합계. 손계산 일치. 시간 포함 날짜는 직접 테스트하지 않음 (논리만 설명, 검증했다고 쓰지 말 것).

### E. average-order-value-by-channel-averageifs-excel-google-sheets
- 제목: How to Find Average Order Value by Sales Channel With AVERAGEIFS in Excel and Google Sheets
- 템플릿: tidy-tabs-average-order-value-by-channel.xlsx (Orders, By channel, How to use)
- 수식: `=AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2)`, `=COUNTIFS(Orders!$B$2:$B$200,A2)`, `=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2)`, 확인 `=C2/B2`.
- 검증: Etsy 3건 $115.50 평균 $38.50, Shopify 3건 $219.50 평균 $73.17 (73.1667), Farmers market 4건 $116.00 평균 $29.00, 전체 10건 $451.00 평균 $45.10. 평균 = 매출/건수 일치. 손계산 일치. 일치 항목 없으면 #DIV/0! 가능 (노트에 적음, 직접 재현은 안 함).

## 2026-10-10 신규 주제 5개 (researcher)
- 데이터: tasks/search-stats.md 없음 (데이터 없음). 3-gate 기준만 적용. 검색량·경쟁도 수치 없음.
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md 와 대조. 이미 쓴 SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/VLOOKUP/MROUND/CEILING/PMT/TEXTJOIN/RANK/WORKDAY 계열 제외.
- 사전 검증: LibreOffice Calc 24.2.7.2 (Linux) 에서 아래 수식을 작은 테스트 파일로 재계산해 값 확인. Excel/Sheets 는 열어 보지 않음. 템플릿 빌드 전 단계이므로 최종 값은 템플릿에서 다시 확인할 것.
- 모든 SERP 1페이지는 Microsoft/Google 공식 도움말이 차지할 가능성이 높음 (확인 안 함, 추정). 그래서 "사업자 상황 + 롱테일"로 비튼다.

순위 (좋은 순)

### 1. sum-visible-rows-only-subtotal-excel-google-sheets
- 제목 예시: How to Sum Only the Visible Rows After Filtering in Excel and Google Sheets
- 검색어: "sum only visible cells after filter excel", "SUM ignores filter", "subtotal filtered rows google sheets"
- 양식: tidy-tabs-subtotal-filtered-sales.xlsx (Sales: Date, Channel, Amount 약 15행 + 자동필터, Totals, How to use). 필터로 Etsy 만 보면 합계가 따라 변함.
- 수식: `=SUBTOTAL(109,C2:C16)` 합계, `=SUBTOTAL(103,C2:C16)` 건수, `=SUM(C2:C16)` 비교, `=SUBTOTAL(101,C2:C16)` 평균.
- 확인: 테스트 파일에서 행 2개 숨긴 뒤 SUBTOTAL(109)=$75.50, SUBTOTAL(103)=3, SUM=$160.50 (이미 확인). 템플릿에서는 LibreOffice 자동필터 적용 후 값 확인. 숨긴 행 처리는 109 vs 9 차이가 앱마다 다를 수 있어 Excel/Sheets 동작은 공식 도움말 링크로만.
- 경쟁: Microsoft SUBTOTAL 도움말, Exceljet, Ablebits. 롱테일 = "Etsy 채널만 필터했을 때 합계".

### 2. rolling-3-month-average-sales-excel-google-sheets
- 제목 예시: How to Add a Rolling 3-Month Average to Your Sales in Excel and Google Sheets
- 검색어: "rolling average excel", "3 month moving average formula", "smooth seasonal sales"
- 양식: tidy-tabs-rolling-average-sales.xlsx (Monthly sales 12개월 계절성 있는 craft/market 매출, Average column, How to use). 12행 + 선택적 선 그래프.
- 수식: `=AVERAGE(B2:B4)` 아래로 채우기, `=IF(COUNT(B2:B4)<3,"",AVERAGE(B2:B4))` 처음 두 달 빈칸, 확인용 `=SUM(B2:B4)/3`.
- 확인: 1200/950/1100/2400/3100/1800 -> 1083.33, 1483.33, 2200.00, 2433.33 (LibreOffice 에서 이미 확인). 손계산 일치.
- 경쟁: Microsoft "moving average" (Analysis ToolPak/차트 추세선), Exceljet, Ablebits. 롱테일 = 계절 매출 평활화.

### 3. prorate-first-month-subscription-excel-google-sheets
- 제목 예시: How to Prorate a Partial Month for a Retainer or Rent in Excel and Google Sheets
- 검색어: "prorated rent formula excel", "prorate monthly fee by days", "partial month charge calculator"
- 양식: tidy-tabs-prorate-partial-month.xlsx (Start dates + monthly fee, Days charged, Prorated amount, How to use). 시작일 10/12/2026 월 $150.
- 수식: 남은 일수 `=EOMONTH(A2,0)-A2+1`, 월 일수 `=DAY(EOMONTH(A2,0))`, 금액 `=ROUND(B2*C2/D2,2)`.
- 확인: 10/12/2026 $150 -> 20일/31일 = $96.77. 02/20/2026 $150 -> 9일/28일 = $48.21 (둘 다 LibreOffice 에서 확인). 일수 칸은 숫자 서식(날짜로 보이는 문제는 troubleshooting 소재).
- 한계: 계약상 일할계산 방식(30일 기준 등)·임대차법 판단은 하지 않음. 산술 예시라고 명시, 계약서를 따르라고 안내.
- 경쟁: 계산기 사이트, Exceljet EOMONTH. 롱테일 = 프리랜서 리테이너/부스 임대료.

### 4. two-way-rate-card-lookup-index-match-excel-google-sheets
- 제목 예시: How to Look Up a Price in a Two-Way Rate Card With INDEX and MATCH in Excel and Google Sheets
- 검색어: "index match match two criteria", "two way lookup excel", "lookup price by row and column"
- 양식: tidy-tabs-two-way-rate-card.xlsx (Rate card: Logo/Brochure/Website x Standard/Rush/Same day, Quote: 드롭다운 2개로 가격 반환, How to use).
- 수식: `=INDEX(B2:D4,MATCH(F1,A2:A4,0),MATCH(G1,B1:D1,0))`, `=IFERROR(...,"Check spelling")`, 비교 `=SUMPRODUCT((A2:A4=F1)*(B1:D1=G1)*B2:D4)`.
- 확인: Brochure + Rush = $500 (확인함), 오타 "Spa" -> "Check spelling". 템플릿에서 9칸 전부 대조. 기존 vlookup-price-list, quantity-price-tiers(1차원)와 다른 2차원 조회. 기존 글의 INDEX/MATCH 한 줄 언급과 링크로 연결.
- 경쟁: Microsoft INDEX/MATCH 도움말, Exceljet, Ablebits (대형). 롱테일 = 서비스 견적 요금표.

### 5. in-cell-bar-chart-rept-excel-google-sheets
- 제목 예시: How to Make a Bar Chart Inside Cells With REPT in Excel and Google Sheets
- 검색어: "in cell bar chart excel", "REPT function bar graph", "sparkline bar google sheets"
- 양식: tidy-tabs-in-cell-bars.xlsx (Product, Units sold, Bar column, How to use). 제품 6개 판매량.
- 수식: `=REPT("█",ROUND(B2/MAX($B$2:$B$7)*20,0))`, 길이 확인 `=LEN(C2)`.
- 확인: 3200/1600/0/450 -> 20/10/0/3 칸 (LibreOffice 에서 확인). 폰트에 따라 블록 문자가 다르게 보일 수 있음(출력 PDF 로 확인). 대안으로 "|" 사용.
- 경쟁: Exceljet, Ablebits, Microsoft 차트 도움말. 롱테일 = 인쇄용 한 장짜리 재고/판매 표.

## 2026-10-10 (참고) 제외한 후보와 이유
- case-pack 올림(CEILING): round-prices-nearest-nickel-99 와 함수 중복.
- 이메일 도메인 추출(FIND/MID): clean-customer-list, split-name 과 유사.


## 2026-10-10 주제 5개 템플릿 빌드와 재계산 결과 (script-writer 준비)
- 환경: LibreOffice Calc 24.2.7.2 (Linux, headless). Excel, Google Sheets 에서는 열어 보지 않음. 샘플 데이터는 모두 가상.
- 빌드: scripts/build-templates.py 에 함수 5개 추가 (subtotal_filtered_sales, rolling_average_sales, prorate_partial_month, two_way_rate_card, in_cell_bars), main 에 등록. 각 시트 US Letter, 한 페이지 폭에 맞춤. PDF 로 용지 크기 letter 확인.
- 1 subtotal: 필터 없음 $200.50 / 15행. Etsy 필터(12행 숨김): SUBTOTAL(109) $75.50, (103) 3, (101) $25.17, SUM 은 $200.50 그대로. Farmers market 필터: $68.50, 6, $11.42. 손계산 일치. 필터 조건과 숨김 행은 openpyxl 로 복사본에 저장해 LibreOffice 로 재계산함 (화살표를 직접 누르지는 않음).
- 2 rolling: C4:C13 = 1083.33, 1483.33, 2200.00, 2433.33, 2183.33, 1616.67, 1766.67, 2183.33, 2983.33, 3633.33. Python 독립 계산과 일치. 차트 없음.
- 3 prorate: 10/12/2026 $150 = 20/31 = $96.77, 02/20/2026 = 9/28 = $48.21, 윤년 02/20/2028 = 10/29 = $51.72 등 7행 모두 일치.
- 4 rate card: 9칸 조합 모두 B5/B6/B7 일치 (Brochure + Rush = $500). 오타 "Spa" -> B5 #N/A, B6 "Check spelling", B7 $0. MATCH 는 대소문자 구분 안 함. 드롭다운 클릭은 테스트하지 않음. 게시글의 F1/G1 은 템플릿에서 Quote!B2/B3.
- 5 bars: 3200/1600/0/450/2400/800 -> 20/10/0/3/15/5 블록. PDF 에서 블록 문자는 렌더되지만 Carlito 에 글리프가 없어 DejaVu Sans 대체 글꼴로 그려지고 문자 사이에 가는 이음선이 보임. 열 E 의 "|" 대안을 함께 제공.
- 상세 셀 주소: tasks/team-2026-10-10-handoff.md 의 Verified values.

## 2026-10-11 신규 주제 5개 (researcher)
- 데이터: tasks/search-stats.md 없음 (데이터 없음). 검색량·경쟁도 수치 없음. 3-gate 기준만 적용.
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md 와 대조. 가까운 기존 글(sum-expenses-by-category, year-to-date, count-orders-by-month, highlight-duplicate-values, invoice-number-generator, loan-payment)과의 차이는 tasks/team-topics-2026-10-11.md 상단에 적음.
- 주제 (좋은 순, 슬러그 앞에 2026-10-11-team-0N- 붙임):
  1. quarterly-sales-totals-sumifs-excel-google-sheets: `=YEAR(A2)&" Q"&ROUNDUP(MONTH(A2)/3,0)` + SUMIFS, 날짜 범위 + EDATE 방식 병행
  2. find-missing-invoice-numbers-countif-excel-google-sheets: `=IF(COUNTIF($A$2:$A$200,D2)=0,"Missing","OK")`
  3. sum-if-cell-contains-text-wildcard-sumif-excel-google-sheets: `=SUMIF(range,"*"&A2&"*",sum_range)`
  4. split-invoice-into-installments-excel-google-sheets: `=ROUND(B1/B2,2)`, 마지막 회차가 남는 센트 흡수, EDATE
  5. marketplace-fee-net-payout-calculator-excel-google-sheets: 수수료율은 가상 입력값, 순수령액과 목표 수령액용 역산 가격
- 세금·법률 판단 없음: 분기는 달력 분기이며 세무 신고 분기 판단을 하지 않음. 누락 번호는 "무효 처리된 인보이스일 수 있음"으로만 안내. 할부는 이자·수수료 없는 산술. 5번 수수료 요율은 실제 마켓 요율이 아니라 사용자가 입력.

## 2026-10-11 주제 5개 템플릿 빌드와 재계산 결과 (template-builder)
- 환경: LibreOffice Calc 24.2 (Linux, headless). Excel, Google Sheets 에서는 열어 보지 않음. 샘플 데이터는 모두 가상. tasks/search-stats.md 없음.
- 빌드: scripts/build-templates.py 에 함수 5개 추가 (quarterly_sales_totals, missing_invoice_numbers, sumif_contains_text, invoice_installments, marketplace_net_payout), main 에 등록. 시트마다 US Letter, 한 페이지 폭에 맞춤, PDF 로 시트당 1쪽 확인.
- 1 quarterly: 2025 Q4 $90.00, 2026 Q1 $405.50, Q2 $630.25, Q3 $600.50, Q4 $779.99, 합계 $2,506.24 (두 방식 일치, Python 일치). 텍스트 날짜는 날짜 범위 방식에서만 빠져 CHECK 표시.
- 2 missing invoice: 1004/1007/1010 누락, 1009 중복, F8 = F4 = 12. 텍스트 숫자 "1005" 와 접두사 "INV-1005" 는 COUNTIF 가 못 찾아 가짜 Missing.
- 3 SUMIF 와일드카드: Uber $68.75 (3건, 대소문자 무시), "ink" 4건 $78.24 ("Pink" 포함), 빈 검색칸은 전체 $216.59, `~*` 리터럴 동작. 텍스트로 입력한 금액은 COUNTIF 에는 잡히고 SUMIF 합계에서는 빠짐.
- 4 installments: $2,500.00/6 = 5 x $416.67 + $416.65, 합계 일치. $1,000.00/3 = 333.33, 333.33, 333.34. 01/31 시작 EDATE 는 02/28, 03/31, 04/30.
- 5 net payout: 6건 손계산 일치, $0.25 상품은 순수령 -$0.02. 목표 $15.00 역산 가격 $16.86 (확인값 $15.01, 센트 반올림 때문). 수수료율 합 100% 면 #DIV/0!.
- 상세 셀 주소: tasks/team-2026-10-11-handoff.md.

## 2026-10-12 신규 주제 5개 (researcher)
- 데이터: tasks/search-stats.md 없음 (데이터 없음). 검색량·경쟁도 수치 없음. 3-gate 기준만 적용.
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md 와 대조. 가까운 기존 글(highlight-duplicate-values, quantity-price-tiers-lookup, year-to-date-sales-total, weighted-average-cost, rank-top-customers)과의 차이는 tasks/team-topics-2026-10-12.md 상단에 적음.
- 주제 (좋은 순, 슬러그 앞에 2026-10-12-team-0N- 붙임):
  1. compare-two-lists-find-missing-countif-excel-google-sheets: `=IF(COUNTIF($D$2:$D$200,A2)=0,"Not found","Found")`, 양방향
  2. tiered-sales-commission-sumproduct-excel-google-sheets: 구간별 요율을 구간 안의 매출에만 적용, SUMPRODUCT
  3. monthly-sales-goal-progress-excel-google-sheets: 달성률, 하루 필요 매출(EOMONTH), 직선 예상치
  4. batch-cost-per-unit-excel-google-sheets: 포장 단가 x 사용량, 판매 가능 개수로 나눈 개당 원가
  5. percent-of-total-product-sales-excel-google-sheets: 구성비, 누적 구성비, 상위 80% 표시
- 세금·법률 판단 없음: 커미션은 서면 계약을 따르라고 안내하는 산술 예시(급여·원천징수 아님). 원가는 재료비만, 인건비·배송·수수료·세금 제외.

## 2026-10-12 주제 5개 템플릿 빌드와 재계산 결과 (template-builder)
- 환경: LibreOffice Calc 24.2 (Linux, headless). Excel, Google Sheets 에서는 열어 보지 않음. 샘플 데이터는 모두 가상. tasks/search-stats.md 없음.
- 빌드: scripts/build-templates.py 에 함수 5개 추가 (compare_two_lists, tiered_commission, sales_goal_progress, batch_cost_per_unit, percent_of_total_sales), main 에 등록. 시트마다 US Letter, 한 페이지 폭에 맞춤, PDF 로 시트당 1쪽 확인. 빌드가 기존 템플릿 .xlsx 도 다시 저장해서 git checkout 으로 되돌리고 새 파일 5개만 남김. openpyxl 은 /usr/bin/python3 에만 있음.
- 1 compare: 10명 vs 8명, A 에만 4건 (ana, maya, omar, jo), B 에만 2건 (chris, pat), 양쪽 6건. 대소문자 무시. 끝 공백은 가짜 Not found, `*` 는 와일드카드로 전부 Found.
- 2 commission: $126.00, $150.00, $560.00, $0.00, $1,640.04, $400.00, 합계 $2,876.04 (두 방식 일치, Python 일치). 음수 매출과 텍스트 매출은 두 방식이 서로 달라짐.
- 3 goal: 10월 합계 $2,699.25 (9/30 매출 제외), 달성률 33.7%, 하루 필요 $278.99, 예상 $6,973.06 Behind. 월말 "Last day", 목표 초과 시 공백, 매출 없는 달 정상.
- 4 batch: 재료비 합계 $151.26, 개당 $6.30, 판매 가능 개당 $6.88. 판매 가능 0개면 공백, 포장당 수량 0 이면 #DIV/0!, 수량 칸 텍스트는 #VALUE!.
- 5 percent: 상위 3개 83.0% (80% 를 넘기는 3번째 포함), 합계 100%. 정렬 안 하면 전부 Top sellers 로 잘못 표시, 매출 비면 #DIV/0!.
- 상세 셀 주소: tasks/team-2026-10-12-handoff.md.

## 2026-10-13 신규 주제 5개 (researcher)
- 데이터: tasks/search-stats.md 없음 (데이터 없음). 검색량·경쟁도 수치 없음. 3-gate 기준만 적용.
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md 와 대조. 가까운 기존 글(receivables-aging-report, year-to-date-sales-total, count-orders-by-month, weighted-average-cost, vlookup-price-list, split-city-state-zip)과의 차이는 tasks/team-topics-2026-10-13.md 상단에 적음.
- 주제 (좋은 순, 슬러그 앞에 2026-10-13-team-0N- 붙임):
  1. invoice-due-date-payment-terms-excel-google-sheets: `=C2+VLOOKUP(D2,$J$2:$K$6,2,FALSE)`, 남은 일수와 Past due 표시
  2. sales-by-day-of-week-sumif-excel-google-sheets: `=TEXT(A2,"ddd")` + SUMIF/AVERAGEIF, 요일별 합계·평균·Best day
  3. repeat-customer-rate-countif-excel-google-sheets: `=COUNTIF(B2:B9,">=2")/COUNTIF(B2:B9,">0")`
  4. cheapest-supplier-cost-per-ounce-index-match-excel-google-sheets: `=(B2+D2)/C2` + INDEX/MATCH/MIN
  5. split-sku-code-into-parts-left-mid-find-excel-google-sheets: LEFT, MID, FIND 로 하이픈 기준 3분할
- 세금·법률 판단 없음: 납기는 날짜 산술일 뿐이며 실제 조건은 계약서를 따르라고 안내. 공급업체 비교는 가격과 배송비만 (최소 주문량·품질 제외). 고객 수치는 주문 건수 비율일 뿐 원인은 설명하지 않음.

## 2026-10-13 주제 5개 템플릿 빌드와 재계산 결과 (template-builder)
- 환경: LibreOffice Calc 24.2 (Linux, headless). Excel, Google Sheets 에서는 열어 보지 않음. 샘플 데이터는 모두 가상. tasks/search-stats.md 없음.
- 빌드: scripts/build-templates.py 에 함수 5개 추가 (invoice_due_dates, sales_by_weekday, repeat_customers, supplier_price_compare, split_sku_code), main 에 등록. 이번에는 importlib 로 새 함수 5개만 실행해서 기존 템플릿 .xlsx 는 건드리지 않음 (git status 에 새 xlsx 5개와 스크립트만). 시트마다 US Letter, 한 페이지 폭에 맞춤, PDF 로 Letter 크기 확인 (Daily sales 탭은 세로로 2쪽). 스크립트 상단 INPUT_FILL 상수를 새로 추가함.
- 주의: 10-12 함수 5개 (compare_two_lists 등)가 현재 scripts/build-templates.py 에 없음. xlsx 는 templates/ 에 있지만 스크립트 전체를 돌려도 그 5개는 재생성되지 않음. 리드가 확인 필요.
- 1 due dates: 8건 모두 Python 날짜 계산과 일치 (INV-1001 09/01 Net 30 = 10/01/2026, -12 Past due), 요약 Open 6 / Past due 2. "Net 20" 은 Check terms, 01/31 + 30일 = 03/02/2026, 소문자 net 30 도 찾음. 텍스트 날짜는 LibreOffice 가 변환해 계산함 (앱 의존).
- 2 weekday: 28일 (09/01~09/28/2026), 요일당 4일. Sat 평균 $159.19, Sun $229.25 (Best day), 총 $2,716.50, Python 일치. 텍스트 날짜는 LibreOffice 가 요일을 붙이지만 F11 CHECK. 빈 시트는 0 / 공백.
- 3 repeat: 13건, 고객 7명 주문, 반복 3명 = 42.9%, 반복 고객 지출 $280.25 (55.8%). 오타 "Ana Ruis" 와 끝 공백은 G7 CHECK, 대소문자는 무시. 주문 없음이면 비율 공백.
- 4 supplier: 최저 Harbor Craft Wholesale $0.2667/oz, 200 oz 비용 $53.33, 최고가와 차이 $53.96. 크기 0 이나 빈 가격은 행이 빠짐, 동점이면 먼저 적힌 공급업체, 가격이 전부 비면 "Enter prices". 텍스트 "$64.00" 은 LibreOffice 가 변환함.
- 5 sku: 10개 코드 모두 str.split 과 일치, 카테고리 합계 51/60/27/68 = 206 OK. 하이픈 1개는 Check SKU, 3개면 마지막 열에 "8OZ-GIFT", 앞 공백은 H8 CHECK, 소문자는 SUMIF 가 무시.
- 상세 셀 주소: tasks/team-2026-10-13-handoff.md. 제목·첫 문단: tasks/hooks-2026-10-13.md.

## 2026-10-14 team-04 last number in a column (template + verify)
- Added last_number_in_column() to scripts/build-templates.py (run alone via exec of the defs; only templates/tidy-tabs-last-number-in-column.xlsx written). Tabs: Log (A date, B balance, C note; 7 entries with blank rows between, rows 2 to 30 yellow), Summary (B2 latest balance, B3 latest date, B4 COUNT, B5/B6 IFERROR versions), How to use. US Letter, fit to width. Fictional data. LibreOffice Calc 24.2 (Linux, headless) only; Excel and Google Sheets not opened.
- Recalculated: B2 $1,402.75, B3 10/02/2026, B4 7, B5/B6 same. Python last numeric value of the 7 balances = 1402.75, last date 10/02/2026: match.
- Edges on copies, recalculated: text note in B14 ignored (still $1,402.75); "$1,500.00" typed as text in B14 with date 10/05/2026 in A14: balance stays $1,402.75 but latest date moves to 10/05/2026; 0 typed last: returns 0, count 8; new row at B20 1500.10: returned, count 8; lower value last (900): 900 returned (sort order does not matter); all cleared: B2 and B3 #N/A, B4 0, B5 and B6 "No entries".

## 2026-10-14 team-05 count-market-saturdays 템플릿 (template-builder)
- 환경: LibreOffice Calc 24.2 (Linux, headless) 에서만 재계산. Excel, Google Sheets 에서는 열어 보지 않음. 샘플 데이터 가상.
- 빌드: scripts/build-templates.py 에 count_weekdays_in_month 추가 (main 호출 목록 끝), 이 함수만 실행. templates/tidy-tabs-count-weekdays-in-month.xlsx (Months, How to use), US Letter, 한 페이지 폭.
- 2026 월별 토요일 5,4,4,4,5,4,4,5,4,5,4,4 (합 52), 일요일 4,4,5,4,5,4,4,5,4,4,5,4 (합 52), 금요일(I3) 5,4,4,4,5,4,5,4,4,5,4,4 모두 Python 일치. I9 OK, 부스비 $25 x 52 = $1,300.00.
- Edge: 토요일에 시작하는 10/03/2026~10/31/2026 = 5 (시작일 포함), 시작일이 끝보다 늦으면 -5, 하루짜리 토요일 1 일요일 0, I3 를 Sun 으로 바꾸면 I4 1111110. 텍스트 날짜 "01/01/2026" 은 LibreOffice 가 변환해 5 를 반환 (앱 의존, 실제 날짜 입력 안내).

## 2026-10-14 team-01 ounces to pounds and ounces (template + verify)
- Environment: LibreOffice Calc 24.2 (Linux, headless recalculation into a scratch dir, read with openpyxl data_only). Not opened in Excel or Google Sheets. Sample data fictional.
- Build: scripts/build-templates.py function ounces_to_pounds_ounces (run alone, other templates untouched) writes templates/tidy-tabs-ounces-to-pounds-ounces.xlsx (Weights, How to use). US Letter, fit to one page wide. Weights B2:B30 yellow inputs; C `=IF(B2="","",INT(B2/16))`, D `=IF(B2="","",MOD(B2,16))`, E label with &; summary H2:H7 (count, ROUND(SUM,2), INT, ROUND(MOD,2), label, CHECK).
- Verified: 10 packages (50, 16, 3.5, 22, 41, 96, 12.25, 9, 135, 28 oz) match Python divmod(oz,16). Total 412.75 oz = 25 lb 12.75 oz, CHECK OK.
- Edge: 0 gives 0 lb 0 oz; 15.99 gives 0 lb 15.99 oz; 0.5 gives 0 lb 0.5 oz; 1700 oz gives 106 lb 4 oz; 32 gives 2 lb 0 oz; 50.1 gives 3 lb 2.1 oz; -5 gives -1 lb 11 oz (negative, not meaningful). Number typed as text '20': converted, 1 lb 4 oz, but SUM skips it so H3 excludes it and H7 shows CHECK. Word 'abc': #VALUE! in C, D, E and H7. All weights cleared: blanks, totals 0, H7 OK. Without ROUND, a total of 2206.34 oz showed leftover 14.3399999999997 in LibreOffice, so the total cells round to 2 decimals.
- Limits: unit conversion only; no carrier rates, weight limits or rounding-up rules.

## 2026-10-14 team-03 listing length check (template-builder)
- Template: templates/tidy-tabs-listing-length-check.xlsx (function listing_length_check in scripts/build-templates.py). Run alone, no other template rewritten. US Letter, fit to width. Sheets: Listings, How to use. LibreOffice Calc 24.2 (Linux, headless) only; not opened in Excel or Google Sheets. Sample data fictional.
- Cells: Listings A2:B11 (10 listings), formulas C2:F30. C `=IF(A2="","",LEN(A2))`; D `=IF(TRIM(B2)="",0,LEN(B2)-LEN(SUBSTITUTE(B2,",",""))+1-(RIGHT(TRIM(B2),1)=","))`; E Too long / OK; F Too many / OK. I2 title limit 60, I3 tag limit 8 (typed samples, yellow); I4 listings 10, I5 too long 1, I6 too many 1, I7 check OK.
- Verified: characters 46, 59, 35, 77, 38, 36, 36, 44, 42, 49 and tags 4, 8, 9, 4, 3, 0, 3, 5, 8, 5 match Python len and str.count.
- Edge: title of exactly 60 chars OK, 61 Too long; 8 tags OK, 9 Too many; empty tags 0 (plain formula without IF returned 1); trailing comma 3 (plain formula returned 4); "a,,b" counts 3; spaces after commas do not change the count; trailing space in a title counts (26 vs 25 after TRIM); limit changed to 46 gives 3 Too long and the 46 char title stays OK; cleared sheet gives 0, 0, 0, OK; overwriting C3 with 5 turns I7 to CHECK.

## 2026-10-14 team-02 convert-time-zones-client-calls 템플릿 (template-builder)
- 환경: LibreOffice Calc 24.2 (Linux, headless) 에서만 재계산. Excel, Google Sheets 에서는 열어 보지 않음. 샘플 데이터 가상.
- 빌드: scripts/build-templates.py 에 time_zone_converter 추가 (main 호출 목록 끝, US_DATETIME 상수와 datetime import 추가), 이 함수만 실행. templates/tidy-tabs-time-zone-converter.xlsx (Calls, Offsets, How to use), US Letter, 폭 맞춤.
- 셀: Calls A2:C9 (고객, 고객 시간대, 내 통화 시간), D 시차 `=IF(B2="","",IFERROR(VLOOKUP(B2,Offsets!$A$2:$B$5,2,FALSE)-VLOOKUP($I$2,Offsets!$A$2:$B$5,2,FALSE),"Check zone"))`, E `=C2+D2/24` (IF 로 감쌈), F Next day / Previous day / Same day. I2 내 시간대(Eastern), I4 다른 날 통화 수, I5 통화 수. Offsets A2:B5 Eastern 0, Central -1, Mountain -2, Pacific -3.
- 8건 모두 Python datetime + timedelta 와 일치 (Eastern 10/14/2026 4:30 PM 은 Pacific 1:30 PM). Previous day 2건 (Cedar Row 10/15 1:00 AM 은 10/14 10:00 PM, Willow Pet 12:15 AM 은 11:15 PM), I4 2, I5 8.
- Edge (복사본 수정 후 재계산): I2 를 Pacific 으로 바꾸면 Eastern 10/14 10:30 PM 이 10/15 1:30 AM Next day, 같은 zone 은 0 시차, 시간 비움은 E, F 공백, 텍스트 시간 "10/14/2026 4:30 PM" 은 E "Check time", 목록에 없는 zone 은 D "Check zone". Eastern 에서 정각 자정 10/14 12:00 AM 은 Pacific 10/13 9:00 PM Previous day.
- 한계: 고정 시간 오프셋만. 서머타임 규칙, 애리조나 대부분과 하와이 등 시간 변경 안 하는 지역, 4개 존 밖은 처리 안 함 (글에서 NIST 링크로 안내).
- 히어로: scripts/make-heroes.py 에 heroes_2026_10_14 (team-02 항목) 추가, 이 항목만 렌더.

## 2026-10-15 리서처 추가 5개 (검색 데이터 없음 - tasks/search-stats.md 없음, 3-gate 기준만 적용, LibreOffice Calc 24.2 에서 확인)
- How to Calculate Average Days to Get Paid in Excel and Google Sheets - `average-days-to-get-paid-excel-google-sheets`
- How to Calculate a 2/10 Net 30 Early Payment Discount in Excel and Google Sheets - `early-payment-discount-2-10-net-30-excel-google-sheets`
- How to Calculate Sell Through Rate in Excel and Google Sheets - `sell-through-rate-inventory-excel-google-sheets`
- How to Calculate How Many Boxes You Need in Excel and Google Sheets - `how-many-boxes-needed-roundup-excel-google-sheets`
- How to Find the Median Order Value in Excel and Google Sheets - `median-order-value-vs-average-excel-google-sheets`

## 2026-10-15 template-builder results (LibreOffice Calc only, 10/14/2026)
Environment: LibreOffice Calc (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with Python math. Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of main(); only these five were run, no older template rewritten. Each template is US Letter, fit to one page wide, yellow input cells, a How to use sheet; PDF export gave one Letter page per tab. Full cell addresses and edge cases are in tasks/team-2026-10-15-handoff.md.
- team-01 templates/tidy-tabs-days-to-get-paid.xlsx: E `=D2-C2`, J4 `=AVERAGE(E2:E30)`. Days 20, 35, 14, 35, 7, 45, 20, 28; average 25.5, fastest 7, slowest 45, 3 over 30 terms days, 2 unpaid, check OK. Edge: unpaid row given a paid date (average 23.1111), none paid (blank results), paid before invoice (-7, not flagged), paid date as text (LibreOffice converted it; app dependent), paid date without invoice date (CHECK).
- team-02 templates/tidy-tabs-early-payment-discount.xlsx: G `=ROUND(D2*$M$2,2)`, F `=C2+$M$3`. Discounts $25.00, $9.60, $44.00, $16.71, $19.20, $6.25, $29.60, $14.51; taken $88.20, missed $26.31, still open $50.36, 3 open. Paid on the deadline day earns the discount, one day later does not. Edge: amount as text (row blank), terms 1/15 net 45 (12.50 on 1,250.00), empty sheet (all 0).
- team-03 templates/tidy-tabs-sell-through-rate.xlsx: E `=C2/B2`. Rates 85.4%, 62.5%, 80.0%, 35.0%, 40.0%, 78.0%, 100.0%, 22.5%; totals 321 of 511 = 62.8%; 2 Slow, 101 units left on them. Edge: received 0 (blank rate), sold above received (Check counts), sold as text (row dropped from totals), sold blank, empty sheet.
- team-04 templates/tidy-tabs-boxes-needed.xlsx: D `=ROUNDUP(B2/C2,0)`. Boxes 5, 4, 1, 5, 4, 6, 0 (total 25, $36.25 at $1.45, 39 empty slots). Edge: units per box 0 or blank (row blank), units as text (row blank), exact multiple 48/12 (4 boxes), decimal 10.5 units, empty sheet.
- team-05 templates/tidy-tabs-median-order-value.xlsx: F6 `=MEDIAN(C2:C100)`. Average $83.35, median $35.25, difference $48.10, largest $640.00, average without orders above $500 $36.96, 1 order above. Edge: even count (median 36.625), blank row, $640.00 as text (skipped, 12 counted), one order, empty list (blank results).

## 2026-10-16 template-builder results (LibreOffice Calc only, 10/15/2026)
Environment: LibreOffice Calc (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with Python math. Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of main(); the five new files are the only changes in templates/ (a full script run rewrites older .xlsx timestamps only, so those were restored with git checkout). US Letter, yellow input cells, a How to use sheet. Full cell addresses and edge cases are in tasks/team-2026-10-16-handoff.md.
- team-01 templates/tidy-tabs-week-ending-date.xlsx: Log D `=A2+MOD($G$2-WEEKDAY(A2,1),7)` (G2 = 6 for Friday), Weeks B `=SUMIFS(...)`. Week endings 10/02, 10/09, 10/16/2026 with 21.5, 17.75, 7.75 hours (47 total, check OK). Edge: Friday stays, Saturday jumps 6 days, Sunday 10/18 to 10/23, 12/30/2026 to 01/01/2027, Sunday-ending week, date as text (CHECK), hours as text (dropped silently).
- team-02 templates/tidy-tabs-free-shipping-threshold.xlsx: D `=IF(AND(C2="Ship",B2>=$I$2),"Free","Pay")`, E `=MAX(0,ROUND($I$2-B2,2))`. 3 Free, 5 Pay, 3 close (within $10.00), 2 Pickup; $50.00 qualifies, $49.99 needs $0.01. Edge: threshold 45 and gap 5, text subtotal, blank subtotal, blank delivery (all CHECK).
- team-03 templates/tidy-tabs-forecast-next-month.xlsx: G2 `=FORECAST(E2,$C$2:$C$25,$A$2:$A$25)`. Sep to Nov 2026 forecasts $1,926.07, $2,009.64, $2,093.21; slope $83.57 per month. Edge: flat (slope 0), falling, 2 months, text or blank cells (pair skipped, check CHECK). Mismatched ranges gave #VALUE! in LibreOffice, not #N/A.
- team-04 templates/tidy-tabs-client-tenure.xlsx: C `=DATEDIF(B2,$I$2,"y")`, D `"ym"`, as of 10/16/2026. Tenures 5 yr 7 mo, 3 yr 11 mo, 2 yr 7 mo, 1 yr 1 mo, 3 yr 0 mo, 7 yr 3 mo, 0 yr 9 mo, 5 yr 9 mo, 0 yr 0 mo, 8 yr 6 mo; 6 clients at 3 years or more. Edge: Feb 29 start, same day, start after as of (Check date instead of #NUM!), text date. Raw DATEDIF #NUM! not run.
- team-05 templates/tidy-tabs-incomplete-rows.xlsx: F `=COUNTBLANK(A2:E2)`. 12 rows: 4 complete, 8 missing something, 2 missing a ZIP. Edge: space-only cell counted as filled, `=""` cell counted as blank (check shows CHECK), all-blank row, empty sheet.

## 2026-10-17 template-builder results (LibreOffice Calc only, 10/16/2026)
Environment: LibreOffice Calc (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with independent Python math. Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of the main block; only these five were run, no older template rewritten. US Letter, yellow input cells, a How to use sheet; PDF export gave Letter pages. Full cell addresses and edge cases are in tasks/team-2026-10-17-handoff.md.
- team-01 templates/tidy-tabs-overlapping-bookings.xlsx: F `=COUNTIFS($B$2:$B$40,B2,$D$2:$D$40,"<="&E2,$E$2:$E$40,">="&D2)-1`, K2 Yes/No switches to < and >. 12 bookings: 9 with a conflict, 12 total; BK-107 shows 2 (overlaps two), identical pair 1 each, same dates on another space OK, end = next start conflict with Yes and OK with No (7 conflicts). Edge: end before start (Check), date as text (that row Check and the other bookings silently miss it), blank row, one day booking with No (own row not subtracted twice). SUMPRODUCT cross-check agrees (K6 OK).
- team-02 templates/tidy-tabs-cash-drawer-count.xlsx: D `=B2*C2` (guarded), F5 `=ROUND(D12-F4,2)`. Counted $309.35 (111 pieces) vs expected $311.85: short $2.50. Three more drawers: Balanced, Short 0.01, Over 0.01 (Decimal agrees). Edge: all blank, float 0, text/negative/fraction count (Check), SUMPRODUCT second total agrees. Unrounded one cent gave 0.00999999999999091. A first draft returned #VALUE! for a text count and was fixed.
- team-03 templates/tidy-tabs-nth-weekday-of-month.xlsx: B8 `=DATE(B2,B3,1)+MOD(B6-WEEKDAY(DATE(B2,B3,1),1),7)+7*(B5-1)` with MONTH check. 2nd Saturday of 11/2026 is 11/14/2026, last Saturday 11/28/2026. 1,260 combinations (2026 to 2028, all months, weekdays, n 1 to 5) and the 12 month schedule match Python calendar with 0 mismatches. Edge: 5th Saturday 11/2026 No such date, Feb 2028 leap, month 13 and text year (Check inputs).
- team-04 templates/tidy-tabs-format-phone-numbers.xlsx: nested SUBSTITUTE, then drop leading 1, then (555) 123-4567 with LEFT, MID, RIGHT. 12 sample numbers: 8 formatted, 4 Check, matches Python regex; 18 more edge cases match (x12, 9 and 11 digits, letters, 5551e34567, +44, slash, blank, number typed as 5551234567). Digit test uses FIND instead of ISNUMBER(--C2) because 5551e34567 converts to a number. Non-breaking space is not stripped (Check).
- team-05 templates/tidy-tabs-scale-recipe-batch.xlsx: D `=ROUND(B6*$C$3,2)`, E ROUNDUP for each. 12 jars to 30 (factor 2.5): 8 ingredients and the lemons 7.5 to 8 match Decimal; factors 0.5, 1, 5, 10, 1.1 also match. Edge: target blank, 0 or text (factor 0, Check target yield), base 0 (blank, Check base yield), text quantity (row skipped, Check quantities).

## 2026-10-18 template-builder results (LibreOffice Calc only, 10/17/2026)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with independent Python math (49 edited copies, 14,154 formula cells, 0 mismatches, no error values). Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of the main block; only these five were run, no older template rewritten. US Letter, yellow input cells, a How to use sheet; PDF export gave one Letter page per tab. Full cell addresses and edge cases are in tasks/team-2026-10-18-handoff.md.
- team-01 templates/tidy-tabs-lookup-two-criteria.xlsx: Lookup B4 `=INDEX(Prices!$C$2:$C$30,MATCH(B1&"|"&B2,Prices!$D$2:$D$30,0))`, B5 SUMIFS, B6 the same with `INDEX(Prices!$A$2:$A$30&"|"&Prices!$B$2:$B$30,0)` and no helper column. All 9 size and finish pairs agree in the three formulas (Large/Gloss $30.00). Edge: missing pair (Not on price list, SUMIFS 0), duplicate pair (MATCH $30.00, SUMIFS $63.00, Duplicate rows), trailing space (no match), lower case (matches), text price (SUMIFS 0, CHECK).
- team-02 templates/tidy-tabs-stock-on-hand-log.xlsx: Stock B2 `=SUMIFS(Log!$D$2:$D$500,Log!$B$2:$B$500,A2,Log!$C$2:$C$500,"Received")`, E2 `=B2-C2+D2`. 30 log rows, 5 SKUs: on hand 6, 3, 11, 15, 9 (44 total = 240 received - 196 sold + 0 adjustments), 2 Reorder, check OK. Edge: SKU with no rows (0), oversold (-15, Check), lower case SKU (counted), trailing space SKU, text qty, blank SKU, unknown type (row Check, check CHECK), text date (no effect).
- team-03 templates/tidy-tabs-email-domain-split.xlsx: B2 `=IFERROR(LEFT(A2,FIND("@",A2)-1),"")`, C2 `=IFERROR(MID(A2,FIND("@",A2)+1,LEN(A2)),"")`, E2 shape test. 14 addresses: 9 OK, 5 Check (two @, no @, space, no dot, trailing space); example.com 3, example.org 2, example.net 2, other 2. 14 more edge addresses match Python. The topic plan's shape formula returned OK for "@example.com", "x@.com" and "y@example.com." in LibreOffice and was tightened. Double dots are not caught.
- team-04 templates/tidy-tabs-on-time-shipping-rate.xlsx: F2 status with `D2<=C2` for On time, Summary B8 `=IF(B4+B5=0,"",B4/(B4+B5))`. 30 orders: 23 on time, 5 late, 2 not shipped, rate 82.1%, average 2.2 days late, worst 4. Edge: shipped before order and text date (Check, check CHECK), all late (0.0%), none shipped (rate blank, no #DIV/0!), blank row.
- team-05 templates/tidy-tabs-price-per-square-foot.xlsx: E6 `=ROUND(B6*C6/144,2)`, F6 `=MAX($B$2,ROUND(E6*D6*$B$1,2))`. 8 jobs at $2.50 per sq ft, $25.00 minimum: subtotal $361.65, 3 at the minimum, 24 x 60 exactly $25.00 shows No, total $376.65 with $15.00 setup. Edge: rates 3.35 and 2.75 (Decimal agrees, check within half a cent per line), qty 0 or blank, text width, fraction qty, rate 0 or text (Check settings), feet typed as inches (warning under 36 sq in).

## 2026-10-19 template-builder results (LibreOffice Calc only, 10/18/2026)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with independent Python math (40 edited copies, 0 mismatches, no error values). Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of the main block; only these five were run (PYTHONDONTWRITEBYTECODE=1), no older template rewritten. US Letter, yellow input cells, a How to use sheet; PDF export gave one Letter page per tab. Conditional formatting (team-01) and data validation (team-02) were checked as rule XML plus helper formulas and the LibreOffice PDF fills; live behavior was not exercised in any spreadsheet UI. Full cell addresses and edge cases are in tasks/team-2026-10-19-handoff.md.
- team-01 templates/tidy-tabs-expiring-dates.xlsx: Items D `=C2-Summary!$B$1`, E status Expired/Soon/OK; rules `$E2="Expired"` and `$E2="Soon"` on A2:E60. As of 10/19/2026, window 30: 3 expired, 9 soon, 8 OK; 0 days and 30 days Soon, 31 days OK, yesterday Expired. Edge: text date (Check), blank date, text as-of date or window (Check settings, no errors).
- team-02 templates/tidy-tabs-validation-rules.xlsx: custom rules `COUNTIF($A$2:$A$200,A2)=1`, `AND(ISNUMBER(B2),B2>=DATE(2026,1,1),B2<=DATE(2026,12,31))`, `IF(ISNUMBER(C2),AND(C2>0,C2=INT(C2)),FALSE)` (stop style, with prompts) read back from the saved XML. Rules check tab repeats them: 12 rows, 9 OK, 3 Fix (INV-1006 twice, qty 2.5). Edge: case, trailing space, text date or qty, year bounds, 3 PM on 12/31/2026. Pasting skips validation (not run).
- team-03 templates/tidy-tabs-sum-with-or.xlsx: B5 SUMIFS+SUMIFS, B6 `SUM(SUMIFS(...,{"Etsy","Shopify"}))`, B7 `SUMPRODUCT(--(either),D)`. Etsy + Shopify $605.50 in all four (14 orders); all 10 channel pairs match Decimal; same channel twice gives $514.00 with SUMIFS+SUMIFS and $257.00 with SUMPRODUCT. Edge: trailing space and text amount (skipped, check still OK), no rows, blank channel (fixed to 7 orders, not 182).
- team-04 templates/tidy-tabs-fill-blanks-above.xlsx: Filled A3 `=IF(blank row,"",IF(Report!A3="",A2,Report!A3))`, B with TRIM. 15 rows, 5 vendors: Pine Hollow Candles $54.00 before, $120.50 after; total $632.00 unchanged. Edge: blank first row, lone space (plain keeps it, TRIM fills it), padded names, numbers, blank separator row.
- team-05 templates/tidy-tabs-mailing-labels.xlsx: 3 x 6 grid on US Letter with `INDEX(Addresses!$A$2:$A$19,(ROW()-1)*3+COLUMN())` joined with CHAR(10). 18 addresses match reading order; ZIP 02108 keeps its zero; 12 addresses leave 6 blanks; a 19th address is not shown and Labels!F5 says CHECK. Not printed on label paper.

## 2026-10-20 리서처 추가 5개 (검색 데이터 없음 - tasks/search-stats.md 없음, 3-gate 기준만 적용, 핵심 수식은 LibreOffice Calc 24.2 에서 확인)
- 상세는 tasks/team-topics-2026-10-20.md. 좋은 순서: 1) months-to-pay-off (NPER) 2) sum across monthly sheets (3D SUM, INDIRECT) 3) outlier order amounts (QUARTILE IQR) 4) ROAS and break-even ROAS 5) Pareto 80/20 cumulative percent.
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md, templates/ 와 대조. 가까운 기존 글(loan-payment-calculator-pmt, break-even-calculator, median-order-value, rank-top-customers, percent-of-total-product-sales, year-to-date-sales-total)과의 차이는 tasks/team-topics-2026-10-20.md 상단에 적음. 최근 며칠의 LEFT/MID/FIND, 단순 SUMIFS, INDEX/MATCH 중심 주제는 피함.
- 3D 참조(Jan:Apr!B2)는 Google Sheets 에서 지원되지 않음: 글에서 명시하고 + 합산 / INDIRECT 대안을 둔다. 테스트는 LibreOffice 만 주장.

## 2026-10-20 template-builder results (LibreOffice Calc only, 10/19/2026)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with independent Python math (106 edited copies plus the 5 unedited files, 12,903 cell comparisons, 0 mismatches; no error values except one intended #VALUE! in team-02). Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of the main block; only these five were run (PYTHONDONTWRITEBYTECODE=1), no older template rewritten. US Letter, yellow input cells, a How to use sheet, a CHECK cell in each; PDF export gave Letter pages (ROAS Campaigns tab landscape). Conditional formatting (team-03, 04, 05) checked as rule XML plus the LibreOffice PDF fills; not seen in a spreadsheet window. Full cell addresses and edge cases are in tasks/team-2026-10-20-handoff.md.
- team-01 templates/tidy-tabs-months-to-pay-off.xlsx: Payoff B5 `=ROUNDUP(NPER(B2/12,-B3,B1),0)` with guards, Schedule tab 60 rows. $5,000.00 at 8 percent, $400.00 a month: 14 months, last payment $38.19 on 12/01/2027, interest $238.19; goal $3,000.00 at $250.00 and 2 percent: 12 months. Edge: 0 percent (13 and an exact 12), payment equal to or below the interest (Payment too low, no #NUM!), over 60 months, payment above the balance (1 month), blank or text inputs. Real limit: the cent-rounded schedule and NPER can differ by a month when the payment is within a cent of balance plus interest (B10 says CHECK).
- team-02 templates/tidy-tabs-sum-across-monthly-sheets.xlsx: B `=Jan!B2+Feb!B2+Mar!B2+Apr!B2`, C `=SUM(Jan:Apr!B2)`, INDIRECT table A9:D20. Sales $14,416.25 in all three. A tab inserted inside Jan:Apr is in the 3D sum only; a misspelled tab shows Check sheet name. Found in LibreOffice: IFERROR(INDIRECT(..)) returned #REF! for a wrong sheet name (the sheet uses IF(ISERROR(..))); "$100" typed as text is converted by + but skipped by SUM; "n/a" gives #VALUE! with +. 3D references are not supported in Google Sheets (not run there).
- team-03 templates/tidy-tabs-outlier-order-amounts.xlsx: Fences B1 `=QUARTILE(Orders!$C$2:$C$100,1)`, B2 Q3, upper B6, lower B7, Orders D flag. 30 orders: Q1 $44.38, Q3 $60.91, fences $19.57 and $85.72, Review on $500.00 and $4.90. A value on the fence is OK, one cent past is Review. Finding: mean + 2 SD is $232.11, so the $500.00 is also caught by that rule; only the $4.90 slips through (lower limit -$100.17). The hook opening must not say the mean and SD rule misses $500.00.
- team-04 templates/tidy-tabs-roas-calculator.xlsx: B2 `=1/B1` (guarded) 2.50 at a 40 percent margin, E ROAS `=D5/B5`, G profit `=D5*$B$1-B5`. 8 campaigns: spend $1,310.00, sales $5,180.00, ROAS 3.95, profit $762.00, one campaign exactly at 2.50 (profit $0.00, Above), 2 Below. Edge: no sales, no spend, text sales, margin 0, text, 40 or 120 percent (Check margin).
- team-05 templates/tidy-tabs-pareto-top-products.xlsx: D cumulative `=SUM($B$2:B2)/SUM($B$2:$B$40)`, E group `=IF(SUM($B$1:B1)/SUM($B$2:$B$40)<Summary!$B$1,"Top","Rest")`. 15 products: 6 Top (40.0 percent of products, 85.3 percent of revenue). Edge: exactly 80.0 percent (Top, next is Rest), crossing 80 percent, unsorted (Sort largest first), refund and text revenue (Check), ties, blanks, one product.

## 2026-10-21 리서처 추가 5개 (검색 데이터 없음 - tasks/search-stats.md 없음, 3-gate 기준만 적용, 핵심 수식은 LibreOffice Calc 24.2 에서 확인)
- 상세는 tasks/team-topics-2026-10-21.md. 좋은 순서: 1) yearly sales growth (RRI, CAGR) 2) count orders by dollar range (COUNTIFS, FREQUENCY) 3) price in effect on a date (LOOKUP / approximate MATCH on dates) 4) batches from materials on hand (MIN, INT, INDEX/MATCH) 5) appointment end times, gaps and buffer (time arithmetic).
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md, templates/ 와 대조. 가장 가까운 기존 글(percent-change-month-over-month, median-order-value, find-outlier-order-amounts, quantity-price-tiers, shipping-weight-tier-lookup, scale-recipe-batch, stock-on-hand, find-overlapping-bookings, time-to-decimal-hours)과의 차이는 tasks/team-topics-2026-10-21.md 상단에 적음. "시간을 15분 단위로 반올림" 주제는 time-to-decimal 글에 MROUND 가 이미 있어 제외.
- FILTER, LET, LAMBDA, XLOOKUP, TODAY, NOW 안 씀. 테스트는 LibreOffice 만 주장 (RRI 는 xlsx 에서 _xlfn.RRI 로 저장해야 LibreOffice 가 읽음 - 템플릿 빌드 때 주의).

## 2026-10-21 template-builder results (LibreOffice Calc only, 10/20/2026)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with independent Python math (Decimal and exact Fraction minutes; 143 edited copies plus 5 re-saved copies and the 5 committed files, 25,329 cell comparisons, 0 mismatches, no error values). Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of the main block; only these five were run (PYTHONDONTWRITEBYTECODE=1), no older template rewritten. US Letter, yellow input cells, a How to use sheet, a CHECK cell in each; PDF export gave Letter pages (Appointment Schedule tab landscape). Conditional formatting checked only as builder rules plus the LibreOffice PDF fills; not seen in a spreadsheet window. RRI is stored as _xlfn.RRI. Full cell addresses, edge cases and the sheet bugs found and fixed are in tasks/team-2026-10-21-handoff.md.
- team-01 templates/tidy-tabs-annual-growth-rate.xlsx: Growth E6 `=_xlfn.RRI(E3,E4,E5)` with guards, E7 `=(E5/E4)^(1/E3)-1`, E10 rebuild check. 2022 to 2026 sales $12,000.00 to $21,500.00, 4 periods: 15.69% (not the 21.456% of the 3 period topic spot check). Yearly changes 18.33%, 11.97%, 20.13%, 12.57% (Check tab). Edge: one period equals simple percent change, flat 0.00%, decline, end 0 (-100.00%), start 0 or negative, reversed or same years, text year or amount (Check inputs, no #NUM!).
- team-02 templates/tidy-tabs-order-size-ranges.xlsx: Ranges D2 `=COUNTIFS(Orders!$C$2:$C$200,">="&A2,Orders!$C$2:$C$200,"<"&B2)`, SUMIFS beside it, FREQUENCY array cross-check C18:C22. 40 orders: 16 / 12 / 9 / 3 in the $0, $25, $50, $100 ranges, $1,633.98 total, check OK. An amount on a limit goes to the higher range. Edge: refund and $0.00 in a Below row, text amount (CHECK), out of order or blank limits, sub-cent amounts (FREQUENCY differs). LibreOffice returned #VALUE! or #N/A from FREQUENCY on an empty list or a blank limit; wrapped in IF/IFERROR.
- team-03 templates/tidy-tabs-price-by-effective-date.xlsx: Orders D2 `LOOKUP(2,1/((Prices!$A$2:$A$40=B2)*(Prices!$B$2:$B$40<=A2)...` with two ISNUMBER tests, IFERROR to No price yet, INDEX and MATCH block at Orders H9:I12. 30 orders: 28 priced, 2 No price yet, total $7,087.50. Change day gets the new price ($43.00 on 02/28/2026, $45.50 on 03/01/2026 for Dark Roast). Edge: unsorted dates (INDEX and MATCH wrong, LOOKUP right, Check order flag), same day twice (lower row wins), trailing space product, text date, blank price (the first draft showed 0 in LibreOffice).
- team-04 templates/tidy-tabs-batches-from-materials.xlsx: Materials D2 `INT(ROUND(B2/C2,10))` with guards, B10 `=MIN(D2:D7)` guarded, B11 INDEX and MATCH. 6 candle materials: 7 batches, limiting material 8 oz jars, 84 candles, to buy for 10 batches: 2.5 lb wax, 5 oz fragrance, 33 jars, 20 boxes. Edge: exactly one batch (leftover 0), short by 0.01, 0 on hand, tie (first listed), blank or 0 need ignored, text or negative quantity (Check inputs). Bare INT(B2/C2) gave the same answers in LibreOffice for six decimal divisions; the ROUND guard was not shown to be needed there.
- team-05 templates/tidy-tabs-appointment-schedule.xlsx: Schedule E2 `=C2+D2/1440`, F2 `=ROUND((C3-E2)*1440,0)`, status vs buffer L1 (15), Slots tab compares whole minutes. 8 groomer appointments 9:00 AM to 6:15 PM: 480 booked minutes, 75 open, 3 Short buffer rows, check OK, 18 slots Taken and 1 Shared. Edge: overlap (negative gap), past midnight (12:30 AM, flagged), blank row in the middle (CHECK), text start or buffer (Check). The "14.999999" floating point effect did not appear with typed times in LibreOffice; ROUND is a precaution (shown only with a start time pasted with extra digits).

## 2026-10-22 리서처 추가 5개 (검색 데이터 없음 - tasks/search-stats.md 없음, 3-gate 기준만 적용, 핵심 수식은 LibreOffice Calc 24.2 에서 확인)
- 상세는 tasks/team-topics-2026-10-22.md. 좋은 순서: 1) compare market events by profit and profit per hour (SUMIFS, INDEX/MATCH MAX) 2) sales by hour of the day (HOUR, COUNTIFS, SUMIFS) 3) convert YYYYMMDD numbers and ISO text to dates (DATE, LEFT, MID) 4) average days between a customer's orders (MAXIFS, MINIFS, COUNTIF) 5) convert foreign currency orders to USD (INDEX/MATCH, ROUND).
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md, templates/ 와 대조. 가장 가까운 기존 글(farmers-market-sales-log, break-even-calculator, sales-by-day-of-week, fix-dates-that-flip, year-to-date-sales-total, last-order-date-per-customer, repeat-customer-rate, average-days-to-get-paid, marketplace-fee-net-payout)과의 차이는 tasks/team-topics-2026-10-22.md 상단에 적음.
- 주의: (a) 시간 범위 SUMIFS(`MOD(A,1)` 비교)는 LibreOffice 에서 10:00:00 행을 빠뜨림(부동소수점) - 글은 HOUR 보조 열을 쓰고 이유를 설명. (b) 월 13 / 2월 30 은 DATE 가 조용히 넘겨서 TEXT 비교 체크 열 필요. (c) MAXIFS/MINIFS 는 xlsx 에 _xlfn. 접두어로 저장해야 LibreOffice 가 읽고, Excel 2019 이상 필요 (글에서 명시). (d) 환율은 독자가 직접 입력, GOOGLEFINANCE 는 안 씀. 세금 판단 없음.
- FILTER, LET, LAMBDA, XLOOKUP, TODAY, NOW 안 씀. 테스트는 LibreOffice 만 주장 (Excel, Google Sheets 에서 실행했다고 쓰지 않음).
## 2026-10-22 template-builder results (LibreOffice Calc only, 10/21/2026)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with independent Python math (Decimal, datetime; 83 edited copies plus 5 re-saved copies and the 5 committed files, 23,874 cell comparisons, 0 mismatches, no error values). Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of the main block; only these five were run (PYTHONDONTWRITEBYTECODE=1), no older template rewritten. US Letter, yellow input cells, a How to use sheet, a CHECK cell in each; PDF export gave Letter pages (Convert and Rates tabs landscape). Conditional formatting and the color scale checked only as builder rules, not seen in a spreadsheet window. MINIFS and MAXIFS are stored as _xlfn.MINIFS and _xlfn.MAXIFS. Full cell addresses, edge cases and the one sheet bug found and fixed are in tasks/team-2026-10-22-handoff.md.
- team-01 templates/tidy-tabs-market-event-comparison.xlsx: Events B2 `=SUMIFS(Sales!$C$2:$C$200,Sales!$B$2:$B$200,A2)`, F2 profit `=B2-C2-D2` (with a text guard), G2 profit per hour with ISNUMBER guard and Check hours, rank I, best-event INDEX and MATCH in L3 and L5, CHECK L8. 6 fictional events, 40 sales, $7,598.00: Weekend Craft Fair best by profit ($1,665.00, 18 h), Downtown Night Market best per hour ($245.00). Edge: 0, blank or text hours (Check hours), misspelled event name (CHECK), event with no sales (negative profit), tie (first listed, shared rank), text fee (Check costs), duplicate event (CHECK).
- team-02 templates/tidy-tabs-sales-by-hour.xlsx: Orders D2 `=IF(A2="","",IF(ISNUMBER(A2),HOUR(A2),"Check time"))`, By hour C2 COUNTIFS and D2 SUMIFS on Orders!$D$2:$D$300, an Outside these hours row, CHECK H7. 50 orders, $1,242.50: busiest hour 10 AM (9 orders, $237.00); 3 orders ($57.00) outside 8 AM to 6 PM including midnight. 10:00:00 AM is hour 10, 9:59:59 AM hour 9. Confirmed again on a probe: a SUMIFS time range on MOD(A,1) gave 30 where the HOUR helper gave 50 (MOD of 10:00:00 AM is 0.416666666656965). Edge: date with no time (hour 0), text time (Check time), duplicate hour (CHECK), order past the last hour.
- team-03 templates/tidy-tabs-yyyymmdd-to-date.xlsx: Convert D2 `=DATE(INT(A2/10000),MOD(INT(A2/100),100),MOD(A2,100))`, E2 `=DATE(LEFT(B2,4),MID(B2,5,2),RIGHT(B2,2))`, F2 ISO `=DATE(LEFT(C2,4),MID(C2,6,2),MID(C2,9,2))+TIME(MID(C2,12,2),MID(C2,15,2),0)`, check columns G, H, I (rebuild with TEXT and compare, exact text with &""), CHECK L4, a By month tab. 60 sample values all OK. Month 13 and Feb 30 roll over silently (02/10/2027, 03/02/2026) and are flagged; leap day 20240229 OK, 20230229 flagged; a trailing space in text gives a wrong date (10/01/2026) that only the check catches; ISO suffixes (Z, offset) ignored, no time zone conversion; VALUE(LEFT()) gave the same date in LibreOffice. Not run in Excel or Sheets.
- team-04 templates/tidy-tabs-days-between-orders.xlsx: Customers B2 COUNTIF, C2 `_xlfn.MINIFS`, D2 `_xlfn.MAXIFS`, E2 `=IF(B2<2,"One order",(D2-C2)/(B2-1))` with Check dates and Not in Orders guards, F2 reorder window `=D2+E2`, CHECK I5. 10 customers, 40 orders: Ann Lopez 49.0 (01/05, 02/02, 04/13/2026), Maple Street Gift Shop 28.0, Noah Chen 119.0, Dana Whitfield One order. Edge: same-day orders (0.0), unsorted rows (unchanged), trailing space in a name (counted as another customer, CHECK), text date (Check dates), customer not in Orders. MIN(IF()) array alternative gave the same dates in LibreOffice (only in the How to use text).
- team-05 templates/tidy-tabs-currency-to-usd.xlsx: Orders E2 `=IFERROR(ROUND(D2*INDEX(Rates!$B$2:$B$10,MATCH(C2,Rates!$A$2:$A$10,0)),2),"Check currency")` core, built with Check amount and Check rate tests (ISNA(MATCH) first), CHECK H5, Rates E:H compare per-row and total rounding. 30 orders, $1,046.40 with fictional sample rates (CAD 0.73, GBP 1.27, EUR 1.08, AUD 0.65, labeled not real); GBP rows differ from the total by $0.01. Edge: unknown, lower case and trailing space codes, blank or text amount, blank, 0, text or negative rate, half cent rounding matches Decimal half-up, duplicate code (CHECK). First draft returned Check rate for an unknown code; fixed. GOOGLEFINANCE not used or tested; no tax advice.

## 2026-10-23 리서처 추가 5개 (검색 데이터 없음 - tasks/search-stats.md 없음, 3-gate 기준만 적용, 핵심 수식은 LibreOffice Calc 24.2 에서 확인)
- 상세는 tasks/team-topics-2026-10-23.md. 좋은 순서: 1) equipment payback period in months (running total, COUNTIF, SUMIF) 2) seasonal sales index by calendar month (AVERAGEIF, AVERAGE) 3) number items within each order (COUNTIF expanding range) 4) split a shared booth fee or bill by sales share (ROUND, SUM, penny fix) 5) how much monthly income varies (STDEV, AVERAGE, SMALL).
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md, templates/ 와 대조. 가장 가까운 기존 글(break-even-calculator, months-to-pay-off-equipment-balance-nper, return-on-ad-spend-roas-break-even, rolling-3-month-average-sales, forecast-next-month-sales, average-annual-sales-growth-cagr, highlight-duplicate-values, stop-duplicate-invoice-numbers, marketplace-fee-net-payout, find-outlier-order-amounts-quartile-iqr)과의 차이는 tasks/team-topics-2026-10-23.md 상단에 적음.
- 주의: (a) payback 의 COUNTIF 는 누적 합이 줄어들지 않는다는 가정(음수 월이 있으면 틀릴 수 있음, 이 경우는 아직 실행 안 함) - 글에서 밝히거나 확인 칸 추가. (b) 계절 지수로 연 목표를 나누면 ROUND 합이 49999.99 처럼 1센트 모자람 - 확인 칸 필요. (c) 품목 번호는 붙어 있는 행만 실행함(떨어진 행, 대소문자, 끝 공백은 아직 안 함). (d) 공동 비용 분할은 남는 1센트를 가장 큰 몫에 더함, 어떻게 나누는 게 공정한지는 판단 안 함. (e) STDEV 사용(STDEV.S 아님)으로 _xlfn 불필요. 세금·감가상각·투자 판단 없음, 샘플은 가상 데이터.
- FILTER, LET, LAMBDA, XLOOKUP, TODAY, NOW 안 씀. 테스트는 LibreOffice 만 주장 (Excel, Google Sheets 에서 실행했다고 쓰지 않음).

## 2026-10-23 template-builder results (LibreOffice Calc only, 10/22/2026)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with independent Python math (Decimal, statistics.stdev; 108 edited copies plus the 5 committed files, 15,648 cell comparisons, 0 mismatches, no error values). Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of the main block; only these five were run (PYTHONDONTWRITEBYTECODE=1), no older template rewritten. US Letter, yellow input cells, a How to use sheet, a CHECK cell in each; PDF export gave Letter pages (Index, Split and Steadiness tabs landscape). Conditional formatting checked only as builder rules, not seen in a spreadsheet window. No _xlfn prefix needed (AVERAGEIF, SUMIF, STDEV, SMALL, SUMPRODUCT, IFERROR, ROUNDUP are all pre-2010 names). Full cell addresses, edge cases and the sheet bugs found and fixed are in tasks/team-2026-10-23-handoff.md.
- team-01 templates/tidy-tabs-equipment-payback.xlsx: Payback B2 cost, B3 `=COUNTIF(C6:C29,"<"&B2)+1` inside text guards, B4 fraction, running total C6 `=SUM(B$6:B6)`, G1 whole months, G4 first month at or above cost (MATCH on column D), CHECK G5. $4,800.00 machine, 24 months of $300.00 to $650.00: payback month 12, 11.25 months (month 11 total $4,660.00, month 12 total $5,220.00). Edge: cost 0, text or blank (message), cost equal to a month total (11 and 11.00), cost 300 (1 and 1.00), cost above the 24 month total (Not paid back in this range), blank month, loss before the payback month (still right), loss after the payback month (count says 3, true first month 2, CHECK), text profit (CHECK).
- team-02 templates/tidy-tabs-seasonal-sales-index.xlsx: Index C2 `=IF(E2=0,"No data",AVERAGEIF(Sales!$B$2:$B$200,A2,Sales!$C$2:$C$200))`, D2 `=C2/AVERAGE(Sales!$C$2:$C$200)` with guards, E years of data, F share of a typed yearly goal (I1), I3 to I6 busiest and slowest, CHECK I11. 24 months (2024, 2025) of a fictional gift shop: Dec index 2.24 (avg $18,350.00), Feb 0.68 ($5,550.00), average month $8,195.83, indexes add up to 12.00; a $90,000.00 goal spread by ROUND gives $89,999.98 (2 cents off, I10). Edge: month with one year of data, month with no data (No data, no #DIV/0!), unequal years (indexes add up to 12.73), month number 13 or text (CHECK), text sales (CHECK), blank or 0 goal, all sales 0 (Check sales), negative month, tie (first listed).
- team-03 templates/tidy-tabs-order-line-numbers.xlsx: Orders C2 `=IF(A2="","",COUNTIF($A$2:A2,A2))`, D2 `=IF(A2="","",COUNTIF($A$2:$A$200,A2))`, E2 Only, First or Last, F2 label such as A1003-2 of 3, summary H1:I8, CHECK I8. 24 items in 9 fictional orders (1 to 4 items each). Edge: non-adjacent rows of one order (numbers and totals right, I6 counts them), a1003 vs A1003 (same order), trailing or leading space (separate order, red cell, CHECK), blank ID with an item (row blank, CHECK), ID with no item (CHECK), 199 random rows, numeric IDs; found in LibreOffice that * and ? in an ID act as wildcards in COUNTIF and that text 007 and number 7 match each other (I7 and I8 flag the first; How to use says the second).
- team-04 templates/tidy-tabs-split-shared-cost.xlsx: Split B1 total, B4:B7 sales, C4 `=ROUND($B$1*B4/SUM($B$4:$B$7),2)` inside a Check sales guard, D4 `=C4+IF(AND(B4=MAX($B$4:$B$7),COUNTIF($B$4:B4,MAX($B$4:$B$7))=1),ROUND($B$1-SUM($C$4:$C$7),2),0)`, E share, F equal split, G difference, CHECK J5. $240.00 booth fee, sales $1,310.00, $890.00, $560.00, $340.00: rounded shares $101.42, $68.90, $43.35, $26.32 add to $239.99, final $101.43 (leftover cent to the largest), $68.90, $43.35, $26.32 add to $240.00. Edge: $100.00 among three equal sellers (33.34, 33.33, 33.33), 1200, 800, 500 (no cent), tie for largest (first row), leftover of -1, +1 and -2 cents, zero sales seller, all sales 0, negative or text sales, text total (Check sales and CHECK, no #DIV/0!), total with a fraction of a cent (86.455 gives CHECK), half cent rounds up.
- team-05 templates/tidy-tabs-income-steadiness.xlsx: Steadiness B2 `=AVERAGE(Income!$B$2:$B$13)` (guarded), B3 `=STDEV(...)`, B4 variation (B3/B2, shown as a percent), B5 lowest, B6 `(SMALL(r,1)+SMALL(r,2)+SMALL(r,3))/3`, B7 `COUNTIF(r,"<"&(B2-B3))`, B8 count, B9 highest, CHECK B10. 12 months (Oct 2025 to Sep 2026): average $2,650.00, STDEV $577.61, variation 0.2180 (21.8%), lowest $1,600.00, 3 lowest average $1,866.67, 2 months below average minus one standard deviation. Edge: identical months (0.0%), 1, 2 and 3 months (messages instead of #DIV/0! and #NUM!), blank month skipped, text entry (CHECK), average 0 (message), negative income, income in cents, no income.

## 2026-10-24 리서처 추가 5개 (검색 데이터 없음 - tasks/search-stats.md 없음, 3-gate 기준만 적용, 핵심 수식은 LibreOffice Calc 24.2 에서 확인)
- 상세는 tasks/team-topics-2026-10-24.md. 좋은 순서: 1) prepaid retainer hours remaining per client (SUMIFS, subtraction) 2) monthly calendar grid from month and year (DATE, WEEKDAY, MONTH) 3) trailing 12 month sales (SUMIFS, EDATE) 4) inventory count variance, physical vs expected (IF, ABS, SUMIF, COUNTIF) 5) weighted score to compare options (SUMPRODUCT, RANK).
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md, templates/ 와 대조. 가장 가까운 기존 글(freelance-hourly-rate-calculator, week-ending-friday-date-weekday, prorate-first-month-subscription, date-of-nth-weekday-in-month, count-market-saturdays-in-month, rolling-3-month-average-sales, year-to-date-sales-total-sumifs, quarterly-sales-totals-sumifs, stock-on-hand-from-in-out-log, cash-drawer-count-over-short, cheapest-supplier-cost-per-ounce)과의 차이는 tasks/team-topics-2026-10-24.md 상단에 적음.
- 주의: (a) 달력 격자는 6개월(10/2026, 02/2026, 08/2026, 05/2027, 03/2026, 02/2028)을 Python calendar 와 대조해 일치, 그 밖의 경우(월 13, 텍스트)는 아직 안 함. (b) TTM 은 날짜가 매월 1일이어야 함(중간 날짜는 `<=` 에서 빠짐), 누락된 달은 개수 칸으로 알림. (c) 재고 차이의 "다른 항목 수"는 빈 칸을 세지 않도록 COUNTIF ">0" + "<0" 사용(아직 안 돌림). (d) 가중 점수는 점수가 독자 본인의 의견이며 어떤 업체를 고르라는 판단은 안 함. (e) 리테이너는 시간만 다루고 요율·청구·계약 판단 없음. 모든 샘플은 가상 데이터.
- FILTER, LET, LAMBDA, XLOOKUP, TODAY, NOW 안 씀 (TTM 은 TODAY 대신 입력한 기준 월). 테스트는 LibreOffice 만 주장 (Excel, Google Sheets 에서 실행했다고 쓰지 않음).

## 2026-10-24 template-builder results (LibreOffice Calc only, 10/23/2026)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless) recalculation of each built file into a scratch dir, read back with openpyxl data_only, compared with independent Python math (Decimal, datetime, a manual calendar model also checked against Python `calendar`; 164 edited copies plus the 5 committed files, 22,814 cell comparisons, 0 mismatches in the final run, no error values). Not opened in Excel or Google Sheets. Sample data fictional. Functions added to scripts/build-templates.py and registered at the end of the main block; only these five were run (PYTHONDONTWRITEBYTECODE=1), no older template rewritten. US Letter, yellow input cells, a How to use sheet, a CHECK cell in each; PDF export gave Letter pages (the main tab of every template is landscape). Conditional fills appeared in the PDF render but were not seen in a spreadsheet window. No _xlfn prefix needed (SUMIFS, COUNTIFS, EDATE, RANK, SUMPRODUCT, WEEKDAY are old names). Full cell addresses, edge cases, flags and the two sheet fixes are in tasks/team-2026-10-24-handoff.md.
- team-01 templates/tidy-tabs-retainer-hours.xlsx: Clients C2 `=IF(A2="","",SUMIFS(Log!$C$2:$C$200,Log!$B$2:$B$200,A2))`, D2 `=B2-C2` inside a guard (the one cell `=B2-SUMIFS(...)` is only in How to use), E percent, F status (Over by N hours, Low at 20% or less left, OK), CHECK I6. 5 fictional clients, 17 log rows, 59 hours: Birch Street Bakery 12.75 of 20 (7.25 left, 63.75%, OK), Lakeview Dental Office 26.50 of 30 (3.50, Low), Pine Hill Yoga 12.25 of 10 (Over by 2.25 hours), Oak Lane Florist 0 used. Edge: client with no rows, hours used up exactly (0 left, Low), bought 0, trailing space in the Log or Clients name (separate name, CHECK), capital letters ignored, text hours (CHECK), typo client (CHECK), duplicate client row (hours counted twice, CHECK), 20% boundary (8 of 10 Low, 7.9 of 10 OK), refund as negative hours, row 200 read and row 201 not.
- team-02 templates/tidy-tabs-monthly-calendar.xlsx: Calendar B1 year, B2 month, B3 first day, grid A6:G11 with the hook formula inside `IF(ISNUMBER($B$3),...,"")`, events per day A14:G19 `=IF(A6="","",COUNTIF(Events!$A$2:$A$200,A6))`, CHECK J5. October 2026 (starts Thursday, 31 days, 5 rows), 13 fictional events, 11 in the month (two on Saturday 10/10). Edge: all 12 months of 2026 and 2027, 02/2028, 02/2024, 02/2100, 1901 and 9999 limits, December to January, month 13, 0, 2.5, text or blank (grid blank, message, CHECK, no error), year as text or 1900, an event with a time part (not counted, CHECK), a text date (not counted, CHECK), no events.
- team-03 templates/tidy-tabs-trailing-12-month-sales.xlsx: Trailing B1 end month, B2 `=SUMIFS(Sales!$B$2:$B$200,Sales!$A$2:$A$200,">"&EDATE(B1,-12),Sales!$A$2:$A$200,"<="&B1)` inside a guard, B3 months found, B4 previous 12, B5 and B6 change, B7 months in the previous 12, CHECK B8. 36 fictional gift shop months 10/01/2023 to 09/01/2026, end 09/01/2026: $67,090.00 (12 months), previous 12 $63,500.00, change $3,590.00 (5.7%). Edge: year boundary, 9 and 11 month windows (message, no change figure), a missing month (count 11), mid-month end date or data date (CHECK), duplicate month (13 found, CHECK), text sales (skipped, CHECK), text date row (no error), unsorted rows, previous 12 total 0, 199 months.
- team-04 templates/tidy-tabs-inventory-count-variance.xlsx: Count E2 `=IF(C2="","",C2-B2)` inside a number test, F2 `=ROUND(E2*D2,2)` inside a guard, G status, H units off helper, summary J2:K9, CHECK K9. 14 fictional gift shop items (13 counted): 10 items off, 34 units off, net value -$82.15, shortages -$102.30, overages $20.15; Postcard set shows Not counted. Edge: blank counted stays blank (not 0), counted 0 is Short, expected 0 with a count above it, negative or text counted, blank or text cost (Check cost, CHECK), numbers in a row with no item name (CHECK), fractional counts, half cent products round up, row 100 read and row 101 not.
- team-05 templates/tidy-tabs-weighted-score-matrix.xlsx: Score weights B3:E3, scores B5:E9, F5 `=ROUND(SUMPRODUCT($B$3:$E$3,B5:E5)/SUM($B$3:$E$3),10)` inside guards, G5 `=RANK(F5,$F$5:$F$9)`, best option J1, CHECK J5. 5 fictional print shops, weights 0.40, 0.20, 0.30, 0.10: Harbor Print Co. 7.60 (rank 1), Bluebird Printing 7.50, Maple Screen Prints 7.40, Pine Street Press 7.00, Quick Tote Shop 6.60. Found in LibreOffice: two options equal on paper got different ranks from floating point noise until the division was wrapped in ROUND(...,10). Edge: weights adding to 100 (same scores and ranks as 0.40, 0.20, 0.30, 0.10) or to 25 (works), all weights 0 or a negative, blank or text weight (Check weights), a tie (same rank, first listed named), blank or text score (Check scores, CHECK), score 11 or 0 (computed, CHECK), blank option name with scores (CHECK).

## 2026-10-25 리서처 추가 5개 (team-01..05, 검색 데이터 없음 - tasks/search-stats.md 없음, 3-gate 기준만 적용, 핵심 수식은 LibreOffice Calc 24.2 에서 확인)
- 중복 확인: posts/ready, posts/published, tasks/topics-seed.md, tasks/team-topics-*.md, templates/ 의 슬러그와 제목 대조. keyword/categoriz/dimensional/gantt/timeline/[h]:mm/crosstab 검색 결과 기존 글 없음. 가장 가까운 글: sum-expenses-by-category-sumif (합계만, 분류는 안 함), quarterly-sales-totals-sumifs, shipping-cost-weight-tier-lookup, weekly timesheet, project tracker. 제목·본문은 hook-writer 에게.
- 사전 확인(LibreOffice 24.2 headless, 임시 xlsx): 키워드 분류 INDEX/MATCH 와 LOOKUP(2,1/...) 둘 다 SHELL OIL 5521 -> Fuel, USPS PO 0245 -> Postage, 키워드 없음 -> Review 로 일치. 부피무게 12x10x8 / 139 -> ROUNDUP 7, 실제 4.2 lb -> 청구 7. 9:30+8:15+10:00 = [h]:mm 에서 27:45, *24 = 27.75, 일반 시간 서식은 03:45 로 보임(재현 필요). 월별 SUMIFS 크로스탭 1월 Fuel 75, 2월 50. 아직 안 돌린 것: 빈 칸/텍스트 입력, 간트 조건부 서식 규칙(템플릿 빌드 때 XML + PDF 로 확인).

### 1) team-01 — categorize-expenses-by-keyword-search-lookup-excel-google-sheets
- 제목 예시: How to Auto-Categorize Expenses From Bank Descriptions in Excel and Google Sheets
- 검색어: "categorize transactions in excel", "auto categorize expenses google sheets", "excel formula if cell contains word return category", "bank statement categories spreadsheet"
- 만들 양식: `templates/tidy-tabs-categorize-expenses-by-keyword.xlsx`. Keywords 탭(키워드 / 카테고리 표, 예: shell=Fuel, usps=Postage, etsy=Fees, staples=Supplies), Expenses 탭(날짜, 가상 은행 설명 25줄, 금액, 자동 카테고리), 맞는 키워드가 없으면 "Review", 카테고리별 합계(SUMIF), Review 줄 수 CHECK. 가상 데이터, US 도시/ZIP 포함 설명.
- 수식: `=IFERROR(INDEX(Keywords!$B$2:$B$40,MATCH(TRUE,INDEX(ISNUMBER(SEARCH(Keywords!$A$2:$A$40,B2)),0),0)),"Review")` (대안 `LOOKUP(2,1/ISNUMBER(SEARCH(...)),...)`), 빈 키워드 칸 처리(빈 키워드는 모든 설명에 맞음, 주의), `SUMIF` 합계, `COUNTIF(...,"Review")`.
- 확인: LibreOffice 로 25줄 재계산 후 Python 의 `in` 검사와 대조, 키워드가 둘 다 맞을 때(표의 위쪽 우선), 대소문자, 빈 키워드, 와일드카드 문자(`*`, `?`, `~`가 SEARCH 에서 와일드카드로 동작하는지; LibreOffice 설정 의존), 빈 설명. Excel/Sheets 는 돌려보지 않음. 세금 공제 가능 여부 판단 없음(카테고리는 독자가 정함).
- 1페이지: Microsoft/Exceljet/Ablebits 의 "if cell contains" 일반 글. "bank description to category" 롱테일로 비튼다.

### 2) team-02 — dimensional-weight-billable-weight-roundup-max-excel-google-sheets
- 제목 예시: How to Calculate Dimensional Weight and Billable Weight for a Package in Excel and Google Sheets
- 검색어: "dimensional weight calculator excel", "dim weight formula spreadsheet", "billable weight vs actual weight", "shipping box volume formula excel"
- 만들 양식: `templates/tidy-tabs-dimensional-weight.xlsx`. 입력(길이, 너비, 높이 인치, 실제 무게 lb, 나누는 수 divisor 기본 139 - 노란 칸, 캐리어 가격표 링크는 출처로만, 값은 독자가 확인), 부피 in³, 부피무게, 청구 무게, 어느 쪽이 적용됐는지 표시. Packages 탭에 가상 12개 상자 목록 + 합계. 요금표 조회는 기존 shipping-cost-weight-tier-lookup 글과 연결만 하고 반복하지 않음.
- 수식: `=ROUNDUP(B1*B2*B3/B5,0)` (부피무게), `=MAX(ROUNDUP(B4,0),B6)` (청구 무게), `=IF(B6>ROUNDUP(B4,0),"Dimensional","Actual")`, 입력 칸 `ISNUMBER` 검사.
- 확인: LibreOffice 에서 12x10x8 / 139 = 6.91 -> 7 (사전 확인함), 정확히 나누어떨어지는 경우(부동소수점), 실제 4.2 lb -> 5, 길이 0 이나 텍스트, divisor 166 으로 바꾸기. Python 으로 12개 상자 독립 계산과 대조. 택배사별 규칙/요금 판단 안 함.
- 1페이지: 캐리어 도움말(UPS/FedEx/USPS), 배송 소프트웨어 블로그. "Excel 에서 두 수 중 큰 값 + ROUNDUP" 롱테일로 비튼다.

### 3) team-03 — simple-gantt-chart-conditional-formatting-excel-google-sheets
- 제목 예시: How to Make a Simple Gantt Chart With Conditional Formatting in Excel and Google Sheets
- 검색어: "simple gantt chart excel template", "gantt chart google sheets conditional formatting", "project timeline spreadsheet", "highlight date range in cells"
- 만들 양식: `templates/tidy-tabs-simple-gantt-chart.xlsx`. 작업 8개(작업명, 시작, 끝, 일수), 첫 열 날짜 머리글 28일(시작 날짜 한 칸에서 +1), 조건부 서식 규칙 하나로 시작~끝 사이 칸 채움, 오늘 줄은 넣지 않거나 입력한 기준 날짜 사용(TODAY 안 씀), 가상 프로젝트(가을 마켓 부스 준비, 날짜 MM/DD/YYYY). 주말 칸 연한 회색.
- 수식: 조건부 서식 `=AND(F$3>=$B4,F$3<=$C4)`, 일수 `=C4-B4+1` 또는 `=NETWORKDAYS(B4,C4)`, 머리글 `=F3+1`, 끝이 시작보다 빠른 줄 표시 CHECK.
- 확인: LibreOffice 로 규칙 수식을 같은 셀들에서 TRUE/FALSE 로 평가(사전 확인함: 11/02~11/06 에서 TRUE), PDF 내보내기로 채워진 칸 모습을 이미지로 확인(US Letter 가로), openpyxl 로 규칙 XML 읽기. Excel/Sheets 에서 규칙 입력 위치는 Microsoft/Google 도움말 링크, 직접 확인했다고 쓰지 않음.
- 1페이지: Smartsheet/TeamGantt/Vertex42 템플릿 페이지(매우 경쟁적). "조건부 서식 수식 하나로" 롱테일로 비튼다. 기존 freelance-project-tracker 와는 날짜 막대 시각화가 다름.

### 4) team-04 — add-up-hours-over-24-hours-h-mm-format-excel-google-sheets
- 제목 예시: Why Your Total Hours Reset After 24 and How to Fix It in Excel and Google Sheets
- 검색어: "excel sum hours over 24", "total time shows wrong excel [h]:mm", "add hours and minutes excel", "google sheets sum duration"
- 만들 양식: `templates/tidy-tabs-total-hours-over-24.xlsx`. 가상 프리랜서 주간 시간(일별 9:30, 8:15, 10:00 등 H:MM) 목록 3주, 잘못된 서식(h:mm 이 03:45)과 맞는 서식([h]:mm 이 27:45)을 나란히, 소수 시간 `*24`, 시급 곱한 금액($), `TEXT(total,"[h]:mm")` 문자열 버전. 기존 time-to-decimal-hours 글(시작/끝 시각 -> 소수 시간)과 달리 "더한 합계가 24 를 넘을 때 표시" 문제.
- 수식: `=SUM(B2:B8)` + 서식 `[h]:mm`, `=SUM(B2:B8)*24`, `=ROUND(SUM(B2:B8)*24*rate,2)`, `=TEXT(B9,"[h]:mm")`, 분 올림 `=INT(...)`/`MOD`.
- 확인: LibreOffice 에서 9:30+8:15+10:00 = 27:45, *24 = 27.75, h:mm 서식은 03:45 (사전 확인함; CSV 에는 서식 적용 후 값이 나옴). 3주 합계를 Python 분 단위 정수 계산과 대조, 분이 60 넘는 입력, 텍스트 시간("9:30" 이 텍스트일 때), 음수 시간(1900 날짜 시스템, 표시 불가) 확인. 서식 코드는 Microsoft/Google 서식 도움말 링크, Sheets 에서 해 봤다고 쓰지 않음(Sheets 는 duration 서식 메뉴가 다름).
- 1페이지: Microsoft 지원 "Add or subtract time", Exceljet, Ablebits. "27:45" 처럼 구체적인 증상 제목으로 비튼다.

### 5) team-05 — expenses-by-category-by-month-two-way-sumifs-excel-google-sheets
- 제목 예시: How to Total Expenses by Category and Month in One Table With SUMIFS in Excel and Google Sheets
- 검색어: "sum by category and month excel", "monthly expense summary table sumifs", "expenses by month and category google sheets", "sumifs date range month"
- 만들 양식: `templates/tidy-tabs-expenses-by-category-and-month.xlsx`. Expenses 탭(날짜, 카테고리, 금액 가상 60줄, 2026년 1~6월), Summary 탭: 행 = 카테고리 5개, 열 = 월 6개(머리글 1일 날짜 `EDATE`), 행/열 합계, 총합이 원본 합계와 같은지 CHECK, 가장 큰 월 표시. 기존 sum-expenses-by-category(한 기준, 합계만)와 달리 두 기준 표.
- 수식: `=SUMIFS(Expenses!$C$2:$C$500,Expenses!$B$2:$B$500,$A2,Expenses!$A$2:$A$500,">="&B$1,Expenses!$A$2:$A$500,"<"&EDATE(B$1,1))`, `=SUM(B2:B6)`, `=INDEX($B$1:$G$1,MATCH(MAX(B8:G8),B8:G8,0))`, CHECK `=ROUND(SUM(...)-SUM(Expenses!C:C),2)=0`.
- 확인: LibreOffice 로 사전 확인(1월 Fuel 75, 2월 Fuel 50), 60줄 전체를 Python 으로 월/카테고리 합산해 30칸 대조, 오타 카테고리(합계 불일치 -> CHECK), 월 범위 밖 날짜, 1일이 아닌 머리글, 날짜가 텍스트인 줄, 빈 줄. 세금 분류/공제 판단 없음.
- 1페이지: Exceljet/Ablebits SUMIFS 날짜 글, 피벗 테이블 도움말. "피벗 없이, 수식 하나를 끌어서 채우기" 롱테일로 비튼다.


## Hooks 2026-10-25
Honesty: the openings below claim no testing. The formulas were checked only in LibreOffice Calc 24.2 (see the 10/25 briefs above); say that in the post's `> Works in` callout, not here.

### 1) team-01 categorize-expenses-by-keyword
- Title: How to Auto-Categorize Expenses From Bank Descriptions in Excel and Google Sheets
- Opening: Put your keywords and categories in a two-column table, then return the category of the first keyword found in each description with `=IFERROR(INDEX(Keywords!$B$2:$B$40,MATCH(TRUE,INDEX(ISNUMBER(SEARCH(Keywords!$A$2:$A$40,B2)),0),0)),"Review")`. A description like "SHELL OIL 5521" becomes Fuel, and anything with no matching keyword shows Review so you can sort it by hand.
- Why it fits: the searcher has a pasted bank column and wants a category next to it, and they get the formula and the "no match" behavior in the first two sentences.

### 2) team-02 dimensional-weight
- Title: How to Calculate Dimensional Weight and Billable Weight in Excel and Google Sheets
- Opening: Divide length x width x height (inches) by your carrier's divisor and round up with `=ROUNDUP(B1*B2*B3/B5,0)`, then bill the larger of that and the rounded-up actual weight with `=MAX(ROUNDUP(B4,0),B6)`. A 12 x 10 x 8 inch box with a divisor of 139 comes to 7 lb, so a 4.2 lb package is billed at 7.
- Why it fits: the searcher wants the dim weight formula and the "which weight do I pay for" rule at once, with a box size they can check against their own.

### 3) team-03 simple-gantt-chart
- Title: How to Make a Simple Gantt Chart With Conditional Formatting in Excel and Google Sheets
- Opening: Build a row of dates across the top, then add one conditional formatting rule with the formula `=AND(F$3>=$B4,F$3<=$C4)` so every cell between a task's start and end date fills in. Apply it to the whole grid once and each task row becomes a bar.
- Why it fits: the searcher wants a Gantt chart without an add-in or chart tool, and the single custom formula is the answer to "how do I highlight a date range in cells".

### 4) team-04 total-hours-over-24
- Title: How to Sum Hours Over 24 in Excel and Google Sheets
- Opening: Format the total cell with the custom number format `[h]:mm`, because plain `h:mm` restarts at zero every 24 hours. 9:30, 8:15 and 10:00 add up to 27:45 with `[h]:mm` but display as 03:45 without the brackets.
- Why it fits: the searcher is looking at a wrong total like 03:45 and needs the format code first, with the exact numbers showing the symptom and the fix.

### 5) team-05 expenses-by-category-and-month
- Title: How to Total Expenses by Category and Month With SUMIFS in Excel and Google Sheets
- Opening: Put categories down the side and the first day of each month across the top, then fill one formula across the whole table: `=SUMIFS(Expenses!$C$2:$C$500,Expenses!$B$2:$B$500,$A2,Expenses!$A$2:$A$500,">="&B$1,Expenses!$A$2:$A$500,"<"&EDATE(B$1,1))`. It adds each category's amounts dated within that month, with no pivot table.
- Why it fits: the searcher wants a category-by-month grid from a plain expense list, and the answer shows both the date-range trick and that no pivot table is needed.

## Template tests 2026-10-25 (LibreOffice Calc only, 10/24/2026)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless). Each file was edited into scratch copies, recalculated with soffice --headless --convert-to xlsx, and read back with openpyxl data_only, then compared with independent Python math (Decimal, datetime, integer minutes). Not opened in Excel or Google Sheets. Sample data fictional. The five functions were added at the end of scripts/build-templates.py and registered in the main block, but only these five were run (no older template rewritten). The committed xlsx files hold formulas only, with no cached values, the same as the earlier templates (all written by openpyxl); every recalculated copy had 0 error values. No TODAY() or NOW(). Yellow inputs, a How to use sheet, US Letter.
- templates/tidy-tabs-categorize-expenses-by-keyword.xlsx: Expenses D2 `=IF(B2="","",IFERROR(INDEX(Keywords!$B$2:$B$40,MATCH(1,INDEX((Keywords!$A$2:$A$40<>"")*(Keywords!$B$2:$B$40<>"")*ISNUMBER(SEARCH(Keywords!$A$2:$A$40,B2)),0),0)),"Review"))` (rows 2-200), keywords Keywords!A2:B40, category totals F2:G7 (SUMIF, F2:F6 typed, F7 Review), total G8, Review count G9, CHECK G10. Sample 25 lines: Fuel 225.02, Postage 73.90, Fees 26.60, Supplies 171.78, Software 91.97, Review 217.95 (4 lines: diner, market booth deposit, CHECK 1042, Dunkin), total 807.22, CHECK OK. Matched Python lowercase `in` search on all 199 rows in 7 edited copies. First keyword in table order wins (shell/exxon above others tested via a duplicate keyword). Blank keyword rows and keyword-without-category rows are skipped, case ignored, blank description gives "" (its amount stays, CHECK shows CHECK), numeric keyword/description works. Wildcards are live in LibreOffice: keyword `*` matched every description, `sh?ll` matched SHELL, `a~*b` matched the literal a*b only. In LibreOffice the same formula WITHOUT the two <>"" guards returned #N/A for every line when a keyword cell was blank (so no claim about Excel behavior). Not tested: trailing space in a keyword (it just must appear in the description), a new category missing from F2:F6 (totals then differ, CHECK).
- templates/tidy-tabs-dimensional-weight.xlsx: Calculator B4 length, B5 width, B6 height, B7 actual lb, B8 divisor 139 (yellow); B10 volume, B11 `=IF(ok,ROUNDUP(ROUND(B4*B5*B6/B8,8),0),"Check inputs")`, B12 `ROUNDUP(ROUND(B7,8),0)`, B13 `MAX(B11,B12)`, B14 Dimensional/Actual. Packages A2:E100 inputs (12 fictional boxes), F dim, G billable, H applies, summary K1:K6 (CHECK K6, divisor shared from Calculator!B8). Verified: 12x10x8, 4.2 lb, 139 -> volume 960, dim 7, actual 5, billable 7, Dimensional (hand math 960/139 = 6.906). Divisor 166 -> 6. Exact division 139x1x1 / 139 -> 1 (no float bump to 2). Dimensions or divisor 0, negative, text or blank -> "Check inputs" in every result cell, no error. Actual 0 allowed (billable = dim). Actual 7 vs dim 7 -> Actual (tie). Fractional sizes 12.5x9.5x7.25 -> 861 in3, dim 7. 12 sample boxes matched Python at divisor 139 (91 lb total billable, 7 boxes Dimensional, 5 Actual) and at 166 (81 lb, 7 Dimensional, 5 Actual). Text row / blank dimension / numbers with no box name on Packages -> Check inputs and K6 CHECK.
- templates/tidy-tabs-simple-gantt-chart.xlsx: Plan B1 chart start 11/02/2026 (yellow), tasks A4:C23 (8 filled, yellow), D `=IF(AND(B4="",C4=""),"",IF(AND(ISNUMBER(B4),ISNUMBER(C4),C4>=B4),C4-B4+1,"Check dates"))`, E `NETWORKDAYS`, date headers F3:AG3 (F3 `=IF(ISNUMBER($B$1),$B$1,"")`, G3 `=IF(ISNUMBER(F3),F3+1,"")`), weekday names row 2, summary AI3:AJ6 (last day AJ4, tasks past chart AJ5, CHECK AJ6 OK). One CF rule on F4:AG23: `AND(ISNUMBER($B4),ISNUMBER($C4),F$3>=$B4,F$3<=$C4)` (dark green, stopIfTrue) plus a weekend gray rule `WEEKDAY(F$3,2)>5`, read back from the xlsx with openpyxl. The rule formula copied into plain cells over all 20x28 cells matched a Python date check with 0 mismatches (also weekend rule); PDF export (US Letter landscape, page 1 is the chart, page 2 How to use) showed the green bars and gray weekends where expected, bars win over weekend gray. Edge: end before start, text dates, start without end -> "Check dates", no bar, CHECK; one-day task = 1 day; task starting before the chart start or ending after the last day is cut at the edge (the visible part is filled); chart start 12/28/2026 crosses into 2027 correctly; B1 text -> header cells blank, no #VALUE!, CHECK. Not checked: how Excel or Google Sheets import the rule (the How to use sheet says to check it after import).
- templates/tidy-tabs-total-hours-over-24.xlsx: Hours B2:B22 (3 weeks x 7 days, time values, format h:mm, yellow), week rows D2:K4 and all-weeks row 5: F `=SUM(B2:B8)` format h:mm (wrong), G same sum format [h]:mm, H `=ROUND(SUM(B2:B8)*24,2)`, I `=IF(ISNUMBER($E$8),ROUND(SUM(B2:B8)*24*$E$8,2),"Check rate")` (rate E8 $45.00 yellow), J `=TEXT(SUM(B2:B8),"[h]:mm")`, K `=ROUND(SUM(B2:B8)*1440,0)` minutes. CHECK G10 (all times, under 24 h, rate number) and G11 (J equals INT(K/60)&":"&TEXT(MOD(K,60),"00")). Verified, shown values from a CSV export with "save cell content as shown": week 1 (9:30 + 8:15 + 10:00) F 3:45, G 27:45, H 27.75, I $1,248.75, J 27:45; week 2 F 19:00, G 43:00, 43.00, $1,935.00; week 3 F 1:00, G 49:00, 49.00, $2,205.00; all F 23:45, G 119:45, 119.75, $5,388.75, 7185 minutes. Plain CSV export without "as shown" prints 27:45:00 in both columns (the format difference only shows in the displayed value). 40 random sheets (random minutes 1 to 1439 in raw day fractions) matched Python integer minutes, the TEXT string and G11 OK every time. Edge: a time typed as text ("8:15") is skipped by SUM (week total drops 8:15) and G10 shows CHECK; 36 h typed as 1.5 is added (55:30) and CHECK; negative -0.25 is added as a subtraction (13:30) and CHECK; 23:59 + three 10:00 -> 63:29; seven 0:59 -> 6:53; rate text -> "Check rate". Negative displayed totals (#### ) not tested; the How to use sheet says so without a test.
- templates/tidy-tabs-expenses-by-category-and-month.xlsx: Summary B1 first month 01/01/2026 (yellow, format mmm yyyy), C1:G1 `=EDATE(B1,1)`, A2:A6 categories (yellow: Fuel, Postage, Supplies, Fees, Software), B2 `=SUMIFS(Expenses!$C$2:$C$500,Expenses!$B$2:$B$500,$A2,Expenses!$A$2:$A$500,">="&DATE(YEAR(B$1),MONTH(B$1),1),Expenses!$A$2:$A$500,"<"&EDATE(DATE(YEAR(B$1),MONTH(B$1),1),1))`, totals H2:H6 and row 7, K1 biggest month (INDEX/MATCH on B7:G7), K2 its total, K3 all Expenses amounts, K4 amount not in table, CHECK K5. Expenses A2:C61 (60 fictional lines, rows 2-500 read). Verified: Jan Fuel $75.00, Feb Fuel $50.00; all 30 cells matched a Python Decimal sum; row totals Fuel 492.00, Postage 397.35, Supplies 502.35, Fees 607.35, Software 712.35, grand total $2,711.40, biggest month 06/2026 $507.15, CHECK OK. Edge (each compared with Python): header typed as 01/15/2026 gives identical numbers (the DATE(YEAR,MONTH,1) wrapper makes the header's day irrelevant); start 11/20/2025 crosses the year; misspelled category, a date outside the six months, or a text date row -> K4 shows the missing dollars and K5 CHECK; amount typed as text -> CHECK; a date with a time (03/31/2026 11:59:59 PM) counted in March; 02/28, 03/01, 06/30 land in the right months; row 500 read and row 501 not; blank row in the middle OK; empty Expenses -> zeros, K1 blank, OK.

## Researcher 2026-10-25 (topics for posts dated 2026-10-26; no search data, 3-gate only)
Full list in tasks/team-topics-2026-10-25.md. Ranked: (1) most common value with MODE/COUNTIF/INDEX-MATCH, (2) column chart with a goal line (chart not yet built), (3) flag malformed email addresses (FIND/SUBSTITUTE/LEN), (4) clickable order links with HYPERLINK, (5) next yearly anniversary and days until it (DATE/YEAR/MONTH/DAY). Formulas for 1, 3, 4, 5 spot-checked in LibreOffice Calc 24.2 headless; HYPERLINK targets not clicked.

## Builder 2026-10-25 (templates for posts dated 2026-10-26)
Built with `most_common_value`, `sales_vs_goal_chart`, `flag_bad_emails`, `order_links`, `next_anniversary` in scripts/build-templates.py (only those five were run; older .xlsx untouched). Recalculated in LibreOffice Calc 24.2.7.2 and compared with Python models: 79 edited copies, 3,400 cell comparisons, 0 mismatches, no error values; US Letter in every PDF page; chart PDF render shows columns plus a dashed goal line. Not tested: Excel, Google Sheets, chart import, clicking links. Details, cell addresses and flags in tasks/team-2026-10-25-handoff.md. Files: tidy-tabs-most-common-value.xlsx, tidy-tabs-sales-vs-goal-chart.xlsx, tidy-tabs-flag-bad-emails.xlsx, tidy-tabs-order-links.xlsx, tidy-tabs-next-anniversary.xlsx.


## 2026-10-28 빌더 결과 (템플릿 5개, 글 날짜 2026-10-28 PT; LibreOffice Calc 24.2.7.2 에서만 확인, Excel/Google Sheets 는 열어보지 않음)
환경: LibreOffice Calc 24.2.7.2 (Linux, headless). 각 파일을 openpyxl 로 입력만 바꾼 복사본 126개(커밋할 샘플 5개 포함)를 만들어 `soffice --headless --convert-to xlsx` 로 재계산하고 `data_only=True` 로 읽어 독립 Python 모델(Fraction/Decimal 센트 반올림, datetime/calendar, divmod)과 44,992 셀 비교, 불일치 0, 오류값(#NAME? 등, Err:) 0. scripts/build-templates.py 에 함수 5개를 `next_anniversary()` 뒤, 메인 블록 끝에 등록했고 이번에는 이 5개만 실행(예전 .xlsx 는 다시 쓰지 않음). 수식에 AGGREGATE, TODAY, NOW, FILTER, LET, XLOOKUP, 동적 배열 없음. 샘플 데이터는 가상, 날짜는 MM/DD/YYYY, 금액은 US 달러. 상세 셀 주소와 예시 표는 tasks/team-2026-10-28-handoff.md.
- loan: `templates/tidy-tabs-loan-payment-schedule.xlsx` (Schedule, How to use). 입력 C1 12,000.00 / C2 7.00% / C3 36 / C4 11/28/2026. 월 납입액 I1 `=IF(C5="OK",IF(C2=0,ROUND(C1/C3,2),ROUND(-PMT(C2/12,C3,C1),2)),"")`, 이자 D8 `=IF(A8="","",IF($C$2=0,0,ROUND(-IPMT($C$2/12,A8,$C$3,$C$1),2)))`, 원금 E8 `=ROUND(C8-D8,2)`, 잔액 F8 `=ROUND($C$1-E8,2)`, 날짜 B8 `=EDATE($C$4,A8-1)`, 마지막 달 C `=전달잔액+이자`(잔액이 정확히 0.00 이 되도록). 검증값: 월 370.53, 1회차 이자 70.00 / 원금 300.53 / 잔액 11,699.47, 마지막 납입 370.38(11/28 기준 10/28/2029), 총이자 1,338.93, 총납입 13,338.93, CHECK I6 OK. 한계: 마지막 납입은 보통 몇 센트 다름(작은 금액+긴 기간이면 더 큼: $1 / 12% / 60개월은 0.02 와 0.17), 이자는 반올림한 잔액이 아니라 IPMT 이론 일정 기준.
- dates: `templates/tidy-tabs-recurring-billing-dates.xlsx` (Dates, How to use). C1 시작일 11/15/2026, C2 드롭다운 "Twice a month" / "Every 14 days", C3 개수(1~26, 샘플 24). 첫 날짜 B7 `=IF(A7="","",IF($C$2="Every 14 days",$C$1,IF(DAY($C$1)<=15,DATE(YEAR($C$1),MONTH($C$1),15),EOMONTH($C$1,0))))`, 다음 B8 `=IF(A8="","",IF($C$2="Every 14 days",B7+14,IF(DAY(B7)=15,EOMONTH(B7,0),DATE(YEAR(B7),MONTH(B7)+1,15))))`, 요일 `TEXT(B7,"ddd")`, 주말 `WEEKDAY(B7,2)>5`(표시만, 평일로 옮기지 않음). 검증: 샘플 11/15, 11/30, 12/15, 12/31, 01/15/2027 ... 마지막 10/31/2027, 주말 7개, CHECK OK. 13개 시작일(01/31, 02/15, 02/28, 12/31, 02/29/2028 포함) x 2방식 전부 일치. 결정: twice 모드에서 시작일이 15일/말일이 아니면 시작일 이후 첫 15일/말일부터 시작.
- closest: `templates/tidy-tabs-closest-match.xlsx` (Sizes, How to use). 표준 사이즈 H7:S7(최대 12칸, 샘플 8 10 12 14 16 18 20 24 30), 주문 A9:B108(샘플 15줄), 차이 보조열 H:S `=IF(AND(ISNUMBER($B9),ISNUMBER(H$7)),ROUND(ABS($B9-H$7),6),"")`, 가장 가까운 C9 `=INDEX($H$7:$S$7,MATCH(MIN($H9:$S9),$H9:$S9,0))`, 차이 D9 `=MIN($H9:$S9)`, 올려서 맞는 사이즈 E9 `=IF(ISNUMBER(C9),IF(COUNTIF($H$7:$S$7,">="&B9)=0,"None",SMALL($H$7:$S$7,COUNTIF($H$7:$S$7,"<"&B9)+1)),"")`(정렬 불필요), 동점 F9 `=IF(ISNUMBER(C9),IF(COUNTIF($H9:$S9,D9)>1,"Tie, first listed used",""),"")`. 요약 D1:D5, CHECK D5. 샘플: 9인치는 8과 10 동점이라 8, 13 -> 12(동점), 21 -> 20, 22 -> 20(동점), 35 -> 30 / None, 24 -> 24. 동점이면 목록에서 왼쪽에 먼저 적은 사이즈(작은 쪽이 아님). 중복 사이즈, 텍스트 숫자는 CHECK.
- inches: `templates/tidy-tabs-inches-to-fractions.xlsx` (Inches, How to use). C1 드롭다운 8/16/32(샘플 16), 입력 A5:A104(샘플 20개). B5 `=MROUND(A5,1/$C$1)`(음수/텍스트는 "Check number"), 분수 텍스트 C5 `=IF(ISNUMBER(B5),IF(B5=0,"0",TRIM(TEXT(B5,"# ?/"&$C$1))),"")`, 정수부 D5 `=INT(B5)`, 분자 E5 `=ROUND((B5-D5)*$C$1,0)`, 약분 F5 (GCD), 피트-인치 G5, 차이 H5 `=ROUND(B5-A5,8)`. 검증: 0.3125 -> 5/16, 2.7 -> 2.6875 / "2 11/16", 5 -> "5", 0.03 -> 0 ("0"), 12.999 -> 13, 0.09375 -> 0.125(정확히 중간은 올림), 11.875 -> "11 14/16" (약분 "11 7/8"). `"# ?/"&$C$1` 로 만든 서식 문자열이 LibreOffice 에서 동작함(CHOOSE 불필요). TEXT 분수 서식의 Excel/Sheets 동작은 확인 안 함.
- batches: `templates/tidy-tabs-batches-and-rotation.xlsx` (Orders, Helpers, Summary, How to use). Orders B1 배치 크기 10, 주문 A5:A204(샘플 45개), 번호 B5 `=IF(A5="","",COUNTA($A$5:A5))`(빈 줄 건너뜀, ROW 기반이 아님), 배치 C5 `=INT((B5-1)/$B$1)+1`, 위치 D5 `=MOD(B5-1,$B$1)+1`, 담당 E5 `=INDEX(Helpers!$A$4:$A$8,MOD(B5-1,COUNTA(Helpers!$A$4:$A$8))+1)&""`. Helpers A4:A8(샘플 Ana, Ben, Cora). Summary B1:B6, 표 D1:E8. 검증: 45개 / 10 -> 5배치, 마지막 5개, 각 도우미 15개, CHECK OK. 빈 줄, 0개, 200개, 도우미 1~5명, 배치 1/7/10 일치. 한계: 도우미 목록 중간에 빈 칸/중복 이름, 배치 크기 0/2.5/텍스트, 205행 이후 입력은 Summary CHECK 로만 잡힘.
- 실행한 가장자리 사례 요약: loan 0%·1회·60회·$500·$250,000·4.25%·01/31 시작·02/29/2028·잘못된 입력 9종 / dates 시작일 13종·1개·26개·잘못된 입력 6종 / closest 동점, 정렬 안 된 목록, 중간 빈 칸, 소수, 목록 밖(작거나 큼), 중복, 텍스트, 100줄 / inches 8·16·32, 정확히 중간, 12.999, 0, 빈 칸, 텍스트, 음수, 잘못된 정밀도 / batches 위 항목. 한계: Excel, Google Sheets 에서 열어보지 않았고 주장하지 않음(`tested_in` 은 LibreOffice Calc 24.2.7.2 만). 대출·세금·이자 공제 조언 없음.


## 2026-10-29 빌더 결과 (LibreOffice Calc 24.2.7.2 only, never Excel/Sheets)
환경: LibreOffice Calc 24.2.7.2 (Linux, headless). scripts/build-templates.py 에 함수 5개(`hide_div_zero_errors`, `totals_off_by_a_cent`, `giveaway_random_winners`, `average_ignore_zeros`, `customer_tiers`)를 `batches_and_rotation()` 뒤, 메인 블록 끝에 등록했고 이번에는 이 5개만 실행(예전 .xlsx 는 다시 쓰지 않음). 각 파일의 입력만 바꾼 복사본을 `soffice --headless --convert-to xlsx` 로 재계산하고 `data_only=True` 로 읽어 독립 Python 모델과 비교: 총 277개 워크북(샘플 5 + 편집본 272, 그 중 200개는 추첨 샘플 재계산), 결정적인 4개 템플릿에서 2,764 셀 비교. 불일치 0 (설명된 4셀 제외: Visits 에 텍스트 "0", 평균 범위 안의 오류값). 의도한 오류 표시 외에 오류값 없음. 상세 셀 주소, 정확한 수식, 예시 표는 tasks/team-2026-10-29-handoff.md. 이 모든 것은 LibreOffice 에서만 확인했고 Excel/Google Sheets 는 열어보지 않았으며 주장하지 않음(`tested_in`: LibreOffice Calc 24.2.7.2). 샘플 데이터는 가상, 금액은 US 달러.
- div0: `templates/tidy-tabs-hide-div-zero-errors.xlsx` (Examples, How to use). 입력 B5:C16 (Orders, Visits 12주). D `=B5/C5`, E `=IFERROR(B5/C5,"")`, F `=IF(C5=0,"",B5/C5)`, G `=IF(AND(ISNUMBER(B5),ISNUMBER(C5)),IF(C5<>0,B5/C5,""),"")`. 샘플: 오류 5개(D6, D8, D9, D15 #DIV/0!, D10 #VALUE!), 비율 7개, 평균 2.8%, CHECK OK. 오타 시연 J12:J14: `=B5/Visits` #NAME?, IFERROR 는 빈칸으로 숨김, IF 는 #NAME? 유지. 결정: G 에 ISNUMBER(B5) 추가. 가장자리: Visits 전부 0/빈칸/텍스트, 음수, 0/0, Orders 빈칸/텍스트, 텍스트 "0"(LibreOffice 가 나눗셈에서 0 으로 취급해 D, F 가 #DIV/0!), `=G9*2` 는 #VALUE!.
- cent: `templates/tidy-tabs-totals-off-by-a-cent.xlsx` (Invoice, How to use). 입력 A5:C24 (품목, 수량, 단가 소수 4자리, 샘플 8줄). D `=B5*C5`, E `=ROUND(D5,2)`, J5 `=SUM(D5:D24)` 55.19, J6 `=SUM(E5:E24)` 55.22, J7 `=ROUND(ROUND(J5,2)-J6,2)` -0.03, CHECK J9 OK. Decimal ROUND_HALF_UP 모델과 17개 워크북 848 셀 일치(반센트 1.005, 8.015, 11.745, 7.875, 20.745, 1.275 모두 올림). 가장자리: 전부 정확한 센트(차이 0), 반센트 다수(-0.04, -0.10), 수량 1,000,000, 음수 수량, 수량 0, 한 줄, 중간 빈 줄, 단가 텍스트(CHECK). 한계: 어떤 반올림이 필요한지는 말하지 않음.
- giveaway: `templates/tidy-tabs-giveaway-random-winners.xlsx` (Entries, Winners, How to use). 입력 Entries A5:B34 (30명), Winners B1 당첨자 수 3. C `=IF(A5="","",RAND())`, D `=IF(A5="","",RANK(C5,$C$5:$C$34)+COUNTIF($C$5:C5,C5)-1)`, Winners B5 `=IF(A5="","",INDEX(Entries!$A$5:$A$34,MATCH(A5,Entries!$D$5:$D$34,0))&"")`. RAND 는 휘발성이라 샘플에 고정된 당첨자가 없음(열 때마다 새 추첨; 200번 재계산에서 서로 다른 3명 조합 196개). 검증은 구조: 200회 모두 순서 1..30 중복 없음, 당첨 3명, 모두 목록 안, CHECK OK, 30명 전원이 한 번 이상 당첨. 가장자리: 당첨자 1/10, 참가자가 당첨자보다 적음(전원 당첨, 안내 문구), 0명(CHECK), 같은 이름 여러 줄(별개 행), 중간 빈 줄, 숫자 이름, 당첨자 수 0/11/2.5/텍스트(메시지+CHECK), 표 아래 입력(CHECK). 값 고정은 복사 후 값만 붙여넣기(앱에서 실행해 확인하지는 않음).
- average: `templates/tidy-tabs-average-ignore-zeros.xlsx` (Daily sales, How to use). 입력 B5:B32 (28일: 0 이 6일, 빈칸 1, 텍스트 "rain" 1, 판매일 20). E5 `=AVERAGE(B5:B32)` 157.29, E6 `=AVERAGEIF(B5:B32,">0")` 204.48, E7 `SUMIF/COUNTIF`, E8 `=IF(COUNTIF(B5:B32,">0")=0,"No days",AVERAGEIF(B5:B32,">0"))`, 일수 E9 20, 0일 E10 6, 음수 E11, 텍스트 E12, 빈칸 E13, CHECK E14 OK. 오류 시연 G5:G9 (#N/A 1개): AVERAGE #N/A, AVERAGEIF(">0") 41.83 (LibreOffice 결과). 결정: 주제 파일의 SUM/COUNTIF 대신 SUMIF, 0 일수는 COUNTIF(범위,0) 로 음수와 분리. 가장자리: 전부 0 (E6 #DIV/0!, E8 "No days"), 전부 빈칸/텍스트, 값 1개, 음수, 숫자로 저장된 텍스트("45" 는 조용히 제외), 텍스트 "0", 범위 안 #N/A 또는 #DIV/0!(E5 오류, CHECK), 무작위 2세트가 statistics.mean 과 일치.
- tiers: `templates/tidy-tabs-customer-tiers.xlsx` (Customers, Cutoffs, How to use). IFS 확인: 파일 XML 에 `_xlfn.IFS(` 로 쓰면 LibreOffice 에서 계산됨, `IFS(` 로 쓰면 #NAME?, TRUE 캐치올이 없으면 #N/A, 텍스트 "abc" 는 `>=1000` 이 참이라 Gold 가 되므로 ISNUMBER 가드. Cutoffs B2 500, B3 1000, 입력 Customers A5:C104 (25명). D `=IFS(C5>=Cutoffs!$B$3,"Gold",C5>=Cutoffs!$B$2,"Silver",TRUE,"Bronze")` (가드 포함), F 열에 같은 결과의 중첩 IF, E `=IF(OR(D5="Silver",D5="Bronze"),ROUND(IF(D5="Silver",Cutoffs!$B$3,Cutoffs!$B$2)-C5,2),"")`. 샘플: Gold 6명 $8,945.75, Silver 8명 $5,811.74, Bronze 11명 $2,367.99, CHECK OK. 가장자리: 정확히 500.00 Silver / 1,000.00 Gold, 499.99 Bronze(남은 0.01), 0, 빈칸, 음수와 텍스트("Check spend"), 컷오프 순서 뒤바뀜/같음/0/텍스트("Check cutoffs"), 컷오프 200/400, 100명 무작위 2세트 모두 IFS 와 중첩 IF 와 모델이 일치. 한계: 라벨만, 할인이나 대우에 대한 조언 없음.


## 2026-10-30 builder results (LibreOffice Calc 24.2.7.2 only, never Excel/Sheets)
Environment: LibreOffice Calc 24.2.7.2 (Linux, headless). Five new functions in scripts/build-templates.py (`billable_time_rounding`, `all_orders_for_customer`, `overnight_shift_hours`, `lapsed_customer_flag`, `year_over_year_sales`), added after `customer_tiers()` and registered at the end of the main block; only these five were run (older .xlsx untouched). Edited copies (inputs only) were recalculated with `soffice --headless --convert-to xlsx`, read back with openpyxl `data_only=True` and compared with independent Python models: 111 workbooks (5 samples + 106 edited copies), 6,810 cell comparisons, 0 mismatches in the final run, no error values anywhere (the tests found 5 bugs that were fixed and re-run, see the handoff). Every PDF page is US Letter. Formulas use no TODAY, NOW, FILTER, LET, LAMBDA, XLOOKUP, SORT, UNIQUE or dynamic arrays; MAXIFS is stored as `_xlfn.MAXIFS`. Sample data fictional, dollars US. Nothing was opened in Excel or Google Sheets and nothing is claimed about them (`tested_in`: LibreOffice Calc 24.2.7.2). Details, cell addresses, exact formulas and example tables: tasks/team-2026-10-30-handoff.md.
- billable: `templates/tidy-tabs-billable-time-rounding.xlsx` (Time log, Settings, How to use). Settings B1 block 15, B2 rate $85.00. F `=ROUND((E5-D5)*1440,0)`, G `=ROUNDUP(F5/Settings!$B$1,0)*Settings!$B$1`, H `=G5/60`, I `=ROUND(H5*Settings!$B$2,2)`, J extra minutes. Sample: 15 entries, 799 actual, 885 billed, 86 extra, $1,253.75, CHECK OK; 45 stays 45, 46 becomes 60, 0 stays 0. 20 workbooks, 907 comparisons. Edge: blocks 1, 6, 30, 60, seconds round to the nearest minute first, start after end or blank end shows Check times and is left out, invalid settings show Check settings, 60 rows, row below the table gives CHECK. Midnight-crossing shifts are not handled here.
- orders: `templates/tidy-tabs-all-orders-for-customer.xlsx` (Orders, Customer view, How to use). Helper F `=IF(AND('Customer view'!$B$1<>"",B5='Customer view'!$B$1),COUNTIF($B$5:B5,'Customer view'!$B$1),"")`, results `=IFERROR(INDEX(Orders!$A$5:$A$104,MATCH(ROWS($A$4:A4),Orders!$F$5:$F$104,0)),"")`, 12 result rows, CHECK D2. Sample: Alma Reyes 12 orders $258.75, OK. 27 workbooks, 1,326 comparisons. Edge: 0, 1, 11, 12 (OK), 13, 14 and 60 matches (CHECK from 13), capital letters ignored, trailing space in the pick finds nothing, wildcard text gives CHECK, blank amount shows 0 in the list. FILTER not used and not tested. Deviation: 100 slots (rows 5 to 104), not 40.
- overnight: `templates/tidy-tabs-overnight-shift-hours.xlsx` (Shifts, How to use; rate on Shifts B1). F `=IF(J5="OK",MAX(0,MOD(D5-C5,1)-N(E5)/1440),"")`, G `=ROUND(F5*24,2)`, H `=ROUND(G5*$B$1,2)`, I plain `=D5-C5` (negative in LibreOffice, e.g. -16:00), J Review text. Sample: 12 shifts, 77.48 hours, $1,394.64, 2 Review rows, 8 negative plain results, CHECK OK. 14 workbooks, 869 comparisons. Edge: end exactly midnight, break equal to or one minute over the shift, start equals end (Review, also with a date part exactly 24 hours apart), seconds, text and blank cells, invalid rate. How a negative time displays in Excel or Sheets is not claimed. No overtime or pay-law statements.
- lapsed: `templates/tidy-tabs-lapsed-customer-flag.xlsx` (Orders, Customers, Settings, How to use). Settings B1 as-of 10/30/2026 (typed), B2 90 days. B `=IF(A5="","",IF(COUNTIFS(Orders!$B$2:$B$200,A5,Orders!$A$2:$A$200,">0")=0,"No orders",MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A5)))`, C days since, D status, E `SUMIF` spend. Sample: 9 Active, 5 Lapsed, 1 No orders, lapsed spend $1,387.91, CHECK OK; 90 days stays Active, 91 is Lapsed. 27 workbooks, 1,932 comparisons. Edge: as-of date moved, cutoffs 0 to 1000, orders after the as-of date, trailing space, duplicate and unlisted names (CHECK), text dates, invalid settings. MAXIFS needs Excel 2019 or newer, not tested in Excel; no older-Excel alternative was built. TODAY() swap not run. Labels only, no contact or consent advice.
- yoy: `templates/tidy-tabs-year-over-year-sales.xlsx` (Orders, Year over year, How to use). B1 2026. C `=SUMIFS(Orders!$B:$B,Orders!$A:$A,">="&B4,Orders!$A:$A,"<"&EDATE(B4,1))` (blank after the latest order date), D with EDATE(B4,-12) and EDATE(B4,-11), F `=IF(D4=0,"n/a",C4/D4-1)`, totals, a same-months row, best and weakest month, CHECK E2. Sample 2026 $3,169.50 vs 2025 $2,785.87 (+13.8%), March shows n/a, October to December blank, CHECK OK. 23 workbooks, 1,776 comparisons. Edge: year 2025, 2024, 2027, month boundaries with times, leap day 02/29/2024, refund, zero base, text date (CHECK), invalid year. Blank means "after the latest order", not zero sales.
- Bugs the tests found and fixed: lapsed, text-date-only customer showed date 12/30/1899 and Lapsed (COUNTIF changed to COUNTIFS with ">0"); lapsed, text as-of date gave #VALUE! in Days since; yoy, text year gave #VALUE! across the table (AND does not short-circuit); yoy, a year with no months shown gave total 0 and -100% (totals now blank); yoy, best month returned January when every change was 0 (now "No month is up").
- Not run: Excel, Google Sheets, TODAY() in the lapsed as-of cell, FILTER, row counts above the slots (60, 100, 50, 199, 500), daylight-saving, a block of exactly 480.
