# Tidy Tabs 설정 (1회)

Tidy Tabs(tidytabs.blogspot.com)에 하루 5개(5am·8am·10am·1pm·5pm PT = 8am·11am·1pm·4pm·8pm ET)
자동 게시하고, 글마다 GPT로 그린 일러스트를 붙이기 위한 설정. Google·OpenAI
로그인이 필요해서 본인이 직접 해야 한다. 기존 두 블로그와 같은 Google Cloud
프로젝트/OAuth 클라이언트를 재사용해도 된다.

## 0. 이 저장소를 공개(public)로 둘지 정하기

AI 이미지·실제 스크린샷·다운로드용 템플릿(.xlsx)은 이 저장소의 `images/`,
`templates/` 에 커밋되고, 글에서는 `https://raw.githubusercontent.com/<저장소>/main/...`
주소로 불러온다. **이 주소는 저장소가 public 일 때만 열린다.**

- 가장 간단: 이 저장소를 public 으로. 비밀값은 전부 GitHub Secrets 에 있어서 코드에는 없다.
- 코드를 비공개로 두고 싶으면: 파일만 올릴 공개 저장소(예: `tidy-tabs-assets`)나
  다른 공개 호스팅을 쓰고, 저장소 Settings → Secrets and variables → Actions →
  **Variables** 탭에 `ASSET_BASE_URL` 을 그 주소로 등록한다 (끝 슬래시 없이).

## 1. Blogger 에서 새 블로그 만들기

blogger.com → 왼쪽 위 블로그 이름 옆 ▼ → **새 블로그** → 이름 `Tidy Tabs`,
주소 `tidytabs` (이미 있으면 `tidytabs-sheets` 등). 만든 뒤:

- 설정 → 기본 → 언어: English (United States), 시간대: (GMT-08:00) Pacific Time.
- 테마: **Contempo 또는 Essential** (Dynamic Views 는 본문을 JS 로 그려서 색인이 늦다).
- 설정 → 메타 태그 → **검색 설명 사용** 켜기. API 로 보낸 search_description 이
  반영이 안 될 수 있어서(half-handy 분석), 글이 올라간 뒤 몇 개는 편집 화면
  오른쪽 "검색 설명"을 직접 확인한다.
- 페이지 3개 만들기: About(누가, 무엇을 어떻게 테스트하는지), Contact, Privacy.

blogId: 대시보드 주소창 `https://www.blogger.com/blog/posts/<숫자>` 의 숫자.

## 2. Blogger API refresh token

기존 블로그 설정 때 쓴 OAuth 클라이언트 그대로. 이 저장소 폴더에서:

```bash
BLOGGER_CLIENT_ID="..." BLOGGER_CLIENT_SECRET="..." node scripts/get-refresh-token.mjs
```

뜨는 URL 을 열고 **Tidy Tabs 블로그를 소유한 Google 계정**으로 로그인 → 동의 →
출력된 refresh token 복사.

## 3. OpenAI API 키 (GPT 이미지 생성)

https://platform.openai.com/api-keys → **Create new secret key**. 결제 수단 등록이
필요하고, 이미지 모델(gpt-image-1)을 쓰려면 조직 인증(Organization verification)을
요구할 수 있다. 글 하나당 1~2장, 1536x1024 medium 품질.

- 모델을 바꾸고 싶으면 Variables 에 `OPENAI_IMAGE_MODEL` (기본 `gpt-image-1`).
- 이미지는 한 번 그려서 커밋하면 다시 그리지 않는다. 다시 그리려면 그 파일을 지우고
  Actions → **Generate AI images** → Run workflow.

## 4. GitHub Secrets 등록

저장소 → Settings → Secrets and variables → Actions → New repository secret:

| Secret | 값 |
| --- | --- |
| `BLOGGER_CLIENT_ID` | OAuth 클라이언트 ID (기존과 같아도 됨) |
| `BLOGGER_CLIENT_SECRET` | OAuth 클라이언트 보안 비밀 |
| `BLOGGER_REFRESH_TOKEN` | 2단계 토큰 |
| `BLOGGER_BLOG_ID` | 1단계 blogId |
| `OPENAI_API_KEY` | 3단계 키 |
| `ANTHROPIC_API_KEY` | (선택) 한국어 번역 Notion 페이지·초안 생성용. 기존 키 재사용 |
| `NOTION_API_KEY` | (선택) 기존 블로그와 같은 값 |
| `NOTION_PARENT_PAGE_ID` | (선택) 기존과 같은 값. 제목 끝에 `-tidy-tabs` 가 붙는다 |
| `THREADS_ACCESS_TOKEN` | Threads 자동 홍보용. docs/threads-setup.md |

Search Console 은 docs/search-console-setup.md.

## 5. 동작 확인

1. Actions → **Generate AI images** → Run workflow → `images/<글>/ai-1.jpg` 가 커밋되는지.
2. Actions → **Auto-publish to Blogger** → Run workflow (force 체크) → posts/ready 의
   첫 글이 올라가고 posts/published 로 옮겨지는지, 이미지·다운로드 링크가 열리는지.

## 동작 방식

- `publish-clock.yml` (게시 시계): 다음 슬롯까지 기다렸다가 `publish.yml` 을 실행하고
  자기 자신을 다시 예약한다. GitHub 예약(cron)이 이 저장소에서 처음 3시간 동안 한 번도
  안 돌아서 만든 것. cron 은 시계가 끊겼을 때 다시 켜는 용도로만 남겨 뒀다.
- `publish.yml`: 슬롯 이후 첫 실행에서 한 번만 게시 (여러 번 실행돼도 중복 게시 없음).
  게시 전에 빠진 AI 이미지를 먼저 그려서 push 한다.
- 글이 가리키는 `images/`·`templates/` 파일이나 AI 이미지가 없으면 그 글은
  **건너뛰고**(로그에 "Holding") 다음 글을 올린다. 사진 없는 글·죽은 다운로드
  링크가 나가지 않게 하려는 것.
- `generate-images.yml`: posts/drafts·posts/ready 에 글이 push 되면 이미지 생성.
- 매일 콘텐츠팀 루틴(Claude)이 다음 하루치 5개를 템플릿·테스트·Threads 문구까지
  만들어 posts/ready/ 에 넣는다. 모자라면 `generate-posts.yml` 이 11:41 UTC 에
  채운다 (템플릿 없는 글). 수동 실행하면 posts/drafts/ 에 초안만 쓴다.
- 검색 설명: docs/search-description.md. Threads: docs/threads-setup.md.
- 템플릿 다시 만들기: `pip install openpyxl && python3 scripts/build-templates.py`.
- 글 수정: posts/published/ 파일을 고치고 Actions → **Update a published Blogger post**.
