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
