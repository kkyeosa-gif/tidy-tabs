# 검색 설명 자동화 (테마에 1번 붙여넣기)

## 왜 코드로 안 되나
Blogger API v3 에는 검색 설명 필드가 없다. 게시 스크립트가 보내는
`searchDescription` 은 버려지고, 대신 쓸 수 있는 `customMetaData` 는 deprecated 라
저장이 안 된다는 보고가 Google 커뮤니티에 올라와 있다
(https://support.google.com/blogger/thread/343506660). 2026-09-25 에 올라간 첫 글도
검색 설명이 비어 있었다.

## 해결: 테마가 글 첫 문단을 검색 설명으로 쓰게 하기
Tidy Tabs 글은 첫 문단이 "답"(1~2문장)이라 그대로 좋은 검색 설명이 된다. 아래 코드는
**검색 설명이 비어 있는 글에서만** 본문 앞 155자를 `<meta name="description">` 으로
넣는다. 손으로 검색 설명을 넣은 글은 그 값을 그대로 쓴다.

1. Blogger → **테마** → **맞춤설정** 옆 ▼ → **HTML 편집**
2. 편집기 안을 클릭하고 Ctrl+F → `</head>` 검색 (한 군데뿐)
3. `</head>` **바로 윗줄**에 아래를 붙여넣기 → 오른쪽 위 **저장**(디스크 아이콘)

```xml
<meta content='max-image-preview:large' name='robots'/>
<b:if cond='data:view.isPost and !data:blog.metaDescription'>
  <b:with value='data:widgets.Blog.first.posts.first' var='tidyPost'>
    <meta expr:content='snippet(data:tidyPost.body, {length: 155, links: false, linebreaks: false})' name='description'/>
  </b:with>
</b:if>
```

첫 줄(`max-image-preview:large`)은 구글 검색·디스커버에서 대표 이미지를 크게 보여줘도
된다는 표시다 (Google 권장, developers.google.com/search/docs/appearance/google-discover).

저장할 때 오류가 나면 테마는 바뀌지 않는다 (Blogger 가 저장을 거부할 뿐). 그 오류
문구를 알려주면 고친다.

## 확인
`node scripts/check-meta.mjs` 가 게시된 글 페이지를 열어 description 이 있는지
본다. 매일 콘텐츠팀 루틴이 이걸 돌려서, 빠진 글이 있으면 보고한다.
