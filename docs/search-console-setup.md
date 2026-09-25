# Search Console 통계 자동 수집 설정

매주 한 번, Tidy Tabs 블로그(tidytabs.blogspot.com)의 최근 28일 검색
성적(어떤 검색어로 들어왔는지, 어떤 글이 검색에 잘 나오는지)을 자동으로
받아와 `tasks/search-stats.md` 에 써두는 기능. `analyst` 에이전트가 이
파일을 읽는다. Google 로그인이 필요한 단계라 본인이 직접 해야 한다.

## 1. Search Console에 블로그 등록

1. https://search.google.com/search-console 접속 (Blogger랑 같은 Google
   계정으로 로그인)
2. 속성 추가 → **URL 접두어** 선택 → `https://tidytabs.blogspot.com/`
   입력 (끝에 슬래시 꼭 포함)
3. 소유권 확인: "HTML 태그" 방식을 고르면 `<meta ...>` 한 줄을 준다. Blogger
   대시보드 → 설정 → 검색 엔진 최적화 → "Google Search Console" 항목에
   그 인증 코드만 붙여넣으면 된다 (전체 태그 말고 content 값만).
4. 등록 직후엔 데이터가 비어 있다. 며칠~한두 주 지나야 검색 데이터가 쌓인다.

## 2. Search Console API 사용 설정 (같은 Google Cloud 프로젝트 재사용)

1. https://console.cloud.google.com → Blogger/Korea 블로그 설정할 때 썼던
   그 프로젝트를 그대로 써도 된다 (프로젝트는 API 키/토큰 발급용 껍데기일
   뿐이라 여러 블로그가 공유해도 서로 영향 없음).
2. "API 및 서비스 > 라이브러리" → **Google Search Console API** 검색 → 사용
   설정 (이미 다른 블로그에서 켜놨다면 또 안 해도 됨).

## 3. Refresh token 발급 (로컬에서 1회 실행)

**주의**: Tidy Tabs 블로그 소유 Google 계정으로 로그인해야 한다 (Korea
블로그와 계정이 다르면 꼭 확인). 이 저장소(tidy-tabs) 폴더 안에서:

```bash
BLOGGER_CLIENT_ID="Tidy Tabs Blogger 설정 때 쓴 클라이언트 ID" \
BLOGGER_CLIENT_SECRET="Tidy Tabs Blogger 설정 때 쓴 클라이언트 보안 비밀" \
SCOPES="https://www.googleapis.com/auth/webmasters.readonly" \
node scripts/get-refresh-token.mjs
```

Windows PowerShell:
```powershell
$env:BLOGGER_CLIENT_ID="..."
$env:BLOGGER_CLIENT_SECRET="..."
$env:SCOPES="https://www.googleapis.com/auth/webmasters.readonly"
node scripts/get-refresh-token.mjs
```

터미널에 뜨는 URL을 열고 로그인/동의하면 새 refresh token이 출력된다.
(Blogger용 토큰과는 다른 토큰이니 헷갈리지 않게 저장.)

## 4. GitHub Secrets 등록 (이 저장소 kkyeosa-gif/tidy-tabs 에)

| Secret 이름 | 값 |
| --- | --- |
| `GSC_REFRESH_TOKEN` | 3단계에서 받은 refresh token |
| `GSC_SITE_URL` | `https://tidytabs.blogspot.com/` (1단계에서 등록한 URL 그대로, 끝 슬래시 포함) |

`BLOGGER_CLIENT_ID` / `BLOGGER_CLIENT_SECRET` 은 이미 등록되어 있으니 그대로 쓴다.

## 5. 동작 확인

Actions 탭 → **"Fetch search stats"** → **Run workflow** 로 바로 테스트 가능.
성공하면 `tasks/search-stats.md` 가 생기거나 갱신된다. 평소엔 매주 월요일에
자동으로 돈다.

등록 직후라 검색 데이터가 아직 없으면 "데이터 없음"으로 나오는 게 정상이다.
