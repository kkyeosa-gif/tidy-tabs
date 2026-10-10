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
