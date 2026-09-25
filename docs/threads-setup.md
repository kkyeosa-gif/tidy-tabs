# Threads 자동 홍보 설정 (새 계정)

블로그 글이 올라가는 순간, 그 글의 `threads:` 문구(콘텐츠팀 threads-writer 가 씀)와
글 링크가 Threads 에 자동으로 올라간다. 주제 태그는 "Excel".

## 1. 새 Threads 계정을 기존 Meta 앱에 연결 (한 번)
기존 블로그 때 만든 Meta 개발자 앱을 그대로 쓴다.

1. https://developers.facebook.com/apps → 기존 앱 → **앱 역할 → 역할** →
   **사람 추가** → **Threads 테스터** → 새 계정 아이디 입력.
2. 휴대폰 Threads 앱(새 계정으로 로그인) → **설정 → 계정 → 웹사이트 권한 → 초대**
   → 수락.
3. 앱 대시보드 → **사용 사례 → Threads API 액세스 → 설정** → 맨 아래
   **사용자 토큰 생성기** → 새 계정 옆 **액세스 토큰 생성** → 로그인·허용 → 토큰 복사.
   권한은 `threads_basic`, `threads_content_publish` 가 있으면 된다.

## 2. GitHub Secrets 등록 (tidy-tabs 저장소)
| Secret | 값 |
| --- | --- |
| `THREADS_ACCESS_TOKEN` | 1-3 에서 복사한 토큰 |

`THREADS_USER_ID` 는 넣지 않아도 된다 (토큰으로 자동 조회). 다른 블로그의
Threads 값을 넣으면 그 계정에 올라가니 절대 섞지 않는다.

## 3. 확인
Actions → **Threads backfill** → Run workflow. 이미 올라간 글(ZIP 코드 글)이
Threads 에 올라가면 성공. 실패하면 로그의 에러 문구를 알려준다.

## 알아둘 것
- 게시 중 Threads 가 실패해도 블로그 글은 그대로 올라가고, Threads backfill 이
  하루 3번(한 번에 최대 2개) 다시 시도한다.
- 새 계정이 짧은 시간에 API 를 많이 부르면 Meta 가 앱을 잠깐 막는다
  ("API access blocked", half-handy 에서 2026-09-25 발생). 그래서 댓글 자동 답장은
  넣지 않았고, backfill 도 천천히 따라잡는다.
- 토큰은 60일 뒤 만료된다. 만료 전에 `THREADS_ACCESS_TOKEN="현재값" node
  scripts/refresh-threads-token.mjs` (half-handy 저장소에 있음) 로 갱신하거나
  1-3 을 다시 한다. 매일 루틴이 만료 에러를 발견하면 알려준다.
