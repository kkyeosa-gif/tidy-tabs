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
