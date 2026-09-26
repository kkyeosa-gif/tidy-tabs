// Shared helpers for the Blogger publish/update scripts.

export function parseFrontmatter(raw) {
  const m = raw.match(/^---\n([\s\S]*?)\n---\n?([\s\S]*)$/);
  if (!m) throw new Error("post file is missing --- frontmatter ---");
  const meta = {};
  for (const line of m[1].split("\n")) {
    if (!line.trim()) continue;
    const idx = line.indexOf(":");
    if (idx === -1) continue;
    meta[line.slice(0, idx).trim()] = line.slice(idx + 1).trim();
  }
  return { meta, body: m[2].trim() };
}

// Published files (AI images, real screenshots, downloadable templates) live
// in this repo and are served from ASSET_BASE_URL. Defaults to the repo's
// raw.githubusercontent.com URL on main, which only works while the repo is
// public; point ASSET_BASE_URL at another public host otherwise.
export function assetBaseUrl() {
  if (process.env.ASSET_BASE_URL) return process.env.ASSET_BASE_URL.replace(/\/+$/, "");
  const repo = process.env.GITHUB_REPOSITORY;
  if (!repo) throw new Error("ASSET_BASE_URL is not set and GITHUB_REPOSITORY is unknown");
  return `https://raw.githubusercontent.com/${repo}/main`;
}

const isAbsolute = (url) => /^(https?:|mailto:|#)/i.test(url);
export function assetUrl(path) {
  return isAbsolute(path) ? path : `${assetBaseUrl()}/${path.replace(/^\.?\//, "")}`;
}

// Repo-relative files a post points at (![img](images/...) and
// [link](templates/...)). The publisher refuses to post while any is missing,
// so a post never goes live with a broken screenshot or download link.
export function localAssetRefs(body) {
  const refs = new Set();
  for (const m of body.matchAll(/!?\[[^\]]*\]\(([^)\s]+)\)/g)) {
    if (!isAbsolute(m[1])) refs.add(m[1].replace(/^\.?\//, ""));
  }
  return [...refs];
}

// Minimal markdown -> HTML blocks: paragraphs, ## / ### headings, "- " and
// "1. " lists, "> " callouts, | pipe | tables, standalone ![alt](src) images,
// plus inline **bold**, *italic*, `code`, [text](url). Returns one HTML
// string per block so embedImages can slot AI images between them.
function inlineHtml(text) {
  const codes = [];
  return text
    .replace(/`([^`]+)`/g, (_, c) => `\u0000${codes.push(c) - 1}\u0000`)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\*\*(.+?)\*\*/g, "<b>$1</b>")
    .replace(/(^|[^*])\*([^*\s][^*]*?)\*/g, "$1<i>$2</i>")
    .replace(/\[(.+?)\]\((.+?)\)/g, (_, t, url) => `<a href="${escapeAttr(assetUrl(url))}">${t}</a>`)
    .replace(/\u0000(\d+)\u0000/g, (_, i) =>
      `<code>${codes[i].replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</code>`
    );
}

const imageHtml = (url, alt, caption) =>
  `<figure class="post-image" style="margin:1.5em 0;"><img src="${escapeAttr(url)}" alt="${escapeAttr(alt)}" ` +
  `style="max-width:100%;height:auto;border:1px solid #ddd;border-radius:6px;" loading="lazy">` +
  (caption ? `<figcaption style="font-size:0.9em;color:#555;">${inlineHtml(caption)}</figcaption>` : "") +
  `</figure>`;

function tableHtml(rows) {
  const cells = (row) => row.replace(/^\||\|$/g, "").split("|").map((c) => c.trim());
  const isRule = (row) => /^\|?[\s:|-]+\|?$/.test(row);
  const header = cells(rows[0]);
  const bodyRows = rows.slice(1).filter((r) => !isRule(r)).map(cells);
  const td = "border:1px solid #ccc;padding:4px 8px;text-align:left;";
  return (
    `<div style="overflow-x:auto;"><table style="border-collapse:collapse;font-size:0.95em;">` +
    `<thead><tr>${header.map((h) => `<th style="${td}background:#f3f3f3;">${inlineHtml(h)}</th>`).join("")}</tr></thead>` +
    `<tbody>${bodyRows.map((r) => `<tr>${r.map((c) => `<td style="${td}">${inlineHtml(c)}</td>`).join("")}</tr>`).join("")}</tbody>` +
    `</table></div>`
  );
}

export function paragraphsToHtml(md) {
  const out = [];
  let para = [];
  let list = null; // { tag: "ul" | "ol", items: [] }
  let table = null; // raw "| a | b |" lines
  let quote = null;
  const flushPara = () => {
    if (para.length) out.push(`<p>${para.map(inlineHtml).join("<br>")}</p>`);
    para = [];
  };
  const flushList = () => {
    if (list) out.push(`<${list.tag}>${list.items.map((i) => `<li>${inlineHtml(i)}</li>`).join("")}</${list.tag}>`);
    list = null;
  };
  const flushTable = () => {
    if (table) out.push(tableHtml(table));
    table = null;
  };
  const flushQuote = () => {
    if (quote)
      out.push(
        `<blockquote style="border-left:4px solid #2e7d32;background:#f1f8f1;margin:1em 0;padding:0.6em 1em;">` +
          `${quote.map(inlineHtml).join("<br>")}</blockquote>`
      );
    quote = null;
  };
  const flushAll = () => {
    flushPara();
    flushList();
    flushTable();
    flushQuote();
  };

  for (const rawLine of md.split("\n")) {
    const line = rawLine.trim();
    const heading = line.match(/^(#{2,3})\s+(.+)$/);
    const bullet = line.match(/^[-*]\s+(.+)$/);
    const numbered = line.match(/^\d+[.)]\s+(.+)$/);
    const image = line.match(/^!\[([^\]]*)\]\(([^)\s]+)\)(?:\s*\*(.+)\*)?$/);
    if (!line) {
      flushAll();
    } else if (heading) {
      flushAll();
      const tag = heading[1].length === 2 ? "h2" : "h3";
      out.push(`<${tag}>${inlineHtml(heading[2])}</${tag}>`);
    } else if (image) {
      flushAll();
      out.push(imageHtml(assetUrl(image[2]), image[1], image[3]));
    } else if (line.startsWith("|")) {
      if (!table) flushAll();
      (table ??= []).push(line);
    } else if (line.startsWith(">")) {
      if (!quote) flushAll();
      (quote ??= []).push(line.replace(/^>\s?/, ""));
    } else if (bullet || numbered) {
      flushPara();
      flushTable();
      flushQuote();
      const tag = bullet ? "ul" : "ol";
      if (list && list.tag !== tag) flushList();
      if (!list) list = { tag, items: [] };
      list.items.push((bullet ?? numbered)[1]);
    } else {
      flushList();
      flushTable();
      flushQuote();
      para.push(line);
    }
  }
  flushAll();
  return out;
}

export function escapeAttr(s) {
  return String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

// ---- AI photo (OpenAI image API) ---------------------------------------
// frontmatter:
//   image_prompts: <one realistic scene>
//   image_alt: <what the photo shows, 8-15 words>
// One photo per post, generated once by scripts/generate-images.mjs and
// committed as images/<post-file>/<title-slug>-photo.jpg (a descriptive file
// name is a small Google Images signal). It's decoration placed lower in the
// post with an "AI-generated photo" caption; the first image of every post is
// a real, informative hero (the template or the real result), which Blogger
// uses as og:image. See docs/style-guide.md "Images".
// STYLE keeps every screen and label unreadable so the photo never passes for
// a real Excel/Sheets screenshot.
export const IMAGE_STYLE =
  "Realistic photograph, natural window light, shot on a DSLR with shallow depth of field, an authentic " +
  "small-business or home-office setting in the United States, true-to-life colors, no filters. Any text on " +
  "screens, paper, labels, or packaging must be blurred or too small to read. No logos, brand names, or " +
  "watermarks, and no readable close-up of a software interface.";

export function postSlug(file) {
  return file.replace(/\.md$/, "");
}

export function slugify(text, max = 60) {
  return String(text).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, max).replace(/-+$/, "");
}

export function parseImagePrompts(meta) {
  const first = (v) => (v ? v.split("|")[0].trim() : "");
  const prompt = first(meta.image_prompts);
  return prompt ? [{ prompt, alt: first(meta.image_alt) || prompt, n: 1, title: meta.title ?? "" }] : [];
}

export function aiImagePath(file, meta) {
  return `images/${postSlug(file)}/${slugify(meta.title || postSlug(file))}-photo.jpg`;
}

export async function generateAiImage(prompt) {
  const res = await fetch("https://api.openai.com/v1/images/generations", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: process.env.OPENAI_IMAGE_MODEL || "gpt-image-1",
      prompt: `${prompt}\n\nStyle: ${IMAGE_STYLE}`,
      size: "1536x1024",
      quality: process.env.OPENAI_IMAGE_QUALITY || "medium",
      output_format: "jpeg",
      output_compression: 85,
      n: 1,
    }),
  });
  if (!res.ok) throw new Error(`image generation failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  const b64 = data.data?.[0]?.b64_json;
  if (!b64) throw new Error(`image generation returned no image: ${JSON.stringify(data).slice(0, 200)}`);
  return Buffer.from(b64, "base64");
}

// The AI photo goes about two thirds of the way down (never in the first
// blocks: the hero screenshot/render leads), captioned as AI-generated.
export function embedImages(htmlBlocks, images) {
  const out = [...htmlBlocks];
  images.forEach((img) => {
    const idx = Math.min(out.length, Math.max(6, Math.round(out.length * 0.66)));
    out.splice(idx, 0, imageHtml(img.url, img.alt, "AI-generated photo"));
  });
  return out;
}

export async function getAccessToken() {
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.BLOGGER_CLIENT_ID,
      client_secret: process.env.BLOGGER_CLIENT_SECRET,
      refresh_token: process.env.BLOGGER_REFRESH_TOKEN,
      grant_type: "refresh_token",
    }),
  });
  if (!res.ok) throw new Error(`token refresh failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.access_token;
}

export async function publishToBlogger({ title, html, labels, searchDescription, accessToken }) {
  const blogId = process.env.BLOGGER_BLOG_ID;
  const res = await fetch(`https://www.googleapis.com/blogger/v3/blogs/${blogId}/posts/`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ title, content: html, labels, searchDescription }),
  });
  if (!res.ok) throw new Error(`blogger publish failed: ${res.status} ${await res.text()}`);
  return res.json();
}

// Falls back to the post's own opening line if the writer didn't set one,
// so Google always has something better than an auto-truncated snippet to
// show searchers. Trimmed to a word boundary near Google's ~155 char cutoff.
export function deriveSearchDescription(meta, body) {
  if (meta.search_description) return meta.search_description.trim();
  const firstParagraph = body.split(/\n\s*\n/)[0].replace(/\s+/g, " ").trim();
  if (firstParagraph.length <= 155) return firstParagraph;
  return firstParagraph.slice(0, 155).replace(/\s+\S*$/, "") + "...";
}

export async function getBloggerPostByUrl(url, accessToken) {
  const blogId = process.env.BLOGGER_BLOG_ID;
  const path = new URL(url).pathname;
  const res = await fetch(
    `https://www.googleapis.com/blogger/v3/blogs/${blogId}/posts/bypath?path=${encodeURIComponent(path)}`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  if (!res.ok) throw new Error(`blogger lookup by url failed: ${res.status} ${await res.text()}`);
  return res.json();
}

export async function updateBloggerPost({ postId, title, html, labels, searchDescription, accessToken }) {
  const blogId = process.env.BLOGGER_BLOG_ID;
  const res = await fetch(`https://www.googleapis.com/blogger/v3/blogs/${blogId}/posts/${postId}`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ title, content: html, labels, searchDescription }),
  });
  if (!res.ok) throw new Error(`blogger update failed: ${res.status} ${await res.text()}`);
  return res.json();
}

// Editor pass over a generated draft: catch spreadsheet claims that are wrong
// (menu paths, formula syntax, version-specific behavior) before anything
// goes out under this blog's name. Returns the corrected file content.
export async function factCheckAndFix(fileContent) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model: "claude-sonnet-5",
      max_tokens: 6000,
      messages: [
        {
          role: "user",
          content:
            "You are a technical editor fact-checking a Google Sheets / Microsoft Excel how-to post before it publishes. " +
            "Check every specific claim: menu paths and button names, keyboard shortcuts, formula syntax and results, " +
            "and anything that differs between Excel for Windows, Excel for Mac, Excel for the web, and Google Sheets. " +
            "Fix anything wrong. If a claim can't be verified, make it less specific rather than inventing detail. " +
            "Remove any first-person claim of hands-on testing (\"I tested\", \"I tried\") unless the post names " +
            "the tested file and version. Keep the tone and the frontmatter fields exactly as they are.\n\n" +
            "Output ONLY the corrected file content in the exact same format as the input (frontmatter between " +
            "--- markers, then the body) — no preamble, no explanation, no notes about what you changed.\n\n" +
            `---INPUT---\n${fileContent}`,
        },
      ],
    }),
  });
  if (!res.ok) throw new Error(`fact-check pass failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  const text = data.content?.find((b) => b.type === "text")?.text ?? "";
  const trimmed = text.trim();
  if (!trimmed.startsWith("---")) throw new Error(`unexpected fact-check response: ${trimmed.slice(0, 200)}`);
  return trimmed;
}

// Uses the Anthropic API (ANTHROPIC_API_KEY is already required for
// scripts/generate-posts.mjs). An earlier version called the free,
// keyless Google Translate endpoint, but Google blocks it as "automated
// queries" from GitHub Actions' shared IP ranges (HTTP 429) — reliably
// broken, not an occasional flake, so it's not usable from CI.
export async function translateToKorean(title, body) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model: "claude-sonnet-5",
      max_tokens: 4096,
      messages: [
        {
          role: "user",
          content:
            "Translate this English blog post into natural, fluent Korean. Conversational tone, not stiff or literal. " +
            "Respond with exactly this format and nothing else (no preamble, no explanation):\n" +
            "TITLE: <translated title>\n" +
            "BODY:\n<translated body, paragraphs separated by a blank line>\n\n" +
            `Title: ${title}\n\nBody:\n${body}`,
        },
      ],
    }),
  });
  if (!res.ok) throw new Error(`translation failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  // content[0] isn't always the text block (a "thinking" block can come
  // first), so find the text block instead of assuming it's first.
  const text = data.content?.find((b) => b.type === "text")?.text ?? "";
  const m = text.match(/TITLE:\s*(.+?)\n+BODY:\s*([\s\S]*)/);
  if (!m) throw new Error(`unexpected translation response: ${text.slice(0, 200)}`);
  return { title: m[1].trim(), body: m[2].trim() };
}

// Accepts a bare ID, a dashed UUID, or a full Notion page URL pasted into
// the secret, and returns just the 32-hex page ID.
export function notionPageId(raw = "") {
  const hex = raw.replace(/-/g, "").match(/[0-9a-f]{32}(?![0-9a-f])/gi);
  if (!hex) throw new Error("NOTION_PARENT_PAGE_ID does not contain a 32-character Notion page ID");
  return hex[hex.length - 1];
}

export async function createNotionPage({ title, body, sourceTitle, sourceUrl }) {
  const richText = (t) => [{ type: "text", text: { content: t.replace(/\*\*(.+?)\*\*/g, "$1").slice(0, 2000) } }];
  const blocks = [];
  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (!line) continue;
    const heading = line.match(/^(#{2,3})\s+(.+)$/);
    const bullet = line.match(/^[-*]\s+(.+)$/);
    const numbered = line.match(/^\d+[.)]\s+(.+)$/);
    if (heading) {
      const type = heading[1].length === 2 ? "heading_2" : "heading_3";
      blocks.push({ object: "block", type, [type]: { rich_text: richText(heading[2]) } });
    } else if (bullet) {
      blocks.push({ object: "block", type: "bulleted_list_item", bulleted_list_item: { rich_text: richText(bullet[1]) } });
    } else if (numbered) {
      blocks.push({ object: "block", type: "numbered_list_item", numbered_list_item: { rich_text: richText(numbered[1]) } });
    } else {
      blocks.push({ object: "block", type: "paragraph", paragraph: { rich_text: richText(line) } });
    }
  }

  const children = [
    {
      object: "block",
      type: "paragraph",
      paragraph: {
        rich_text: [{ type: "text", text: { content: `원문: ${sourceTitle}`, link: { url: sourceUrl } } }],
      },
    },
    { object: "block", type: "divider", divider: {} },
    ...blocks,
  ].slice(0, 100); // Notion caps children per create request at 100

  const res = await fetch("https://api.notion.com/v1/pages", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.NOTION_API_KEY}`,
      "Notion-Version": "2022-06-28",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      parent: { page_id: notionPageId(process.env.NOTION_PARENT_PAGE_ID) },
      properties: { title: { title: [{ text: { content: title } }] } },
      children,
    }),
  });
  if (!res.ok) throw new Error(`notion page create failed: ${res.status} ${await res.text()}`);
  return res.json();
}

// ---- Threads ---------------------------------------------------------
// The content team writes each post's Threads copy ahead of time in the
// frontmatter `threads:` field (one line, "\n" for line breaks). If it's
// missing, generateThreadsPost writes one at publish time with the same
// rules as .claude/agents/threads-writer.md.
export function threadsCopyFromMeta(meta) {
  return meta.threads ? meta.threads.replace(/\\n/g, "\n").trim() : "";
}

export async function generateThreadsPost(title, body) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json",
    },
    body: JSON.stringify({
      model: "claude-sonnet-5",
      max_tokens: 512,
      messages: [
        {
          role: "user",
          content:
            "Write one Threads post that promotes the spreadsheet how-to below to US small business owners and " +
            "freelancers. Threads rewards posts that read like a real person talking, so: lead with the problem " +
            "in the reader's own words or the one-line fix itself (the reader should get value even without " +
            "clicking), then one concrete detail from the post (a formula, a setting, a before/after number). " +
            "Casual, a little dry humor is fine. Short lines. At most one question. No hashtags, no emoji, no " +
            "'link in bio', no 'Read more', no em dashes, no hype words. Every claim must come from the post; " +
            "never claim hands-on testing the post doesn't mention. Under 380 characters. Do not include a URL; " +
            "code appends it.\n\nRespond with only the post text.\n\n" +
            `Title: ${title}\n\nPost:\n${body}`,
        },
      ],
    }),
  });
  if (!res.ok) throw new Error(`Threads post generation failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  const text = data.content?.find((b) => b.type === "text")?.text ?? "";
  if (!text.trim()) throw new Error("Threads post generation returned empty text");
  return text.trim();
}

// Threads caps a post at 500 characters. Trim the copy, never the URL.
export function buildThreadsText(copy, url) {
  const suffix = `\n\n${url}`;
  const max = 500 - suffix.length;
  const trimmed = copy.length <= max ? copy : copy.slice(0, max).replace(/\s+\S*$/, "");
  return `${trimmed}${suffix}`;
}

// THREADS_USER_ID is optional: the token already identifies the account.
export async function threadsUserId(accessToken) {
  if (process.env.THREADS_USER_ID) return process.env.THREADS_USER_ID;
  const res = await fetch(`https://graph.threads.net/v1.0/me?fields=id,username&access_token=${encodeURIComponent(accessToken)}`);
  if (!res.ok) throw new Error(`threads /me failed: ${res.status} ${await res.text()}`);
  return (await res.json()).id;
}

// Container-then-publish, waiting for the container to finish processing
// (publishing immediately can fail with "Media Not Found"; half-handy hit
// this). Returns { id, permalink }.
export async function publishToThreads({ text, accessToken, topicTag, linkUrl }) {
  const userId = await threadsUserId(accessToken);
  const params = { media_type: "TEXT", text, access_token: accessToken };
  // Link preview card (text-only posts only, per the Threads API docs): shows
  // the post's og:image, i.e. the hero, under the copy.
  if (linkUrl) params.link_attachment = linkUrl;
  if (topicTag) params.topic_tag = topicTag;
  const createRes = await fetch(`https://graph.threads.net/v1.0/${userId}/threads`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params),
  });
  if (!createRes.ok) throw new Error(`threads container create failed: ${createRes.status} ${await createRes.text()}`);
  const { id: creationId } = await createRes.json();

  for (let i = 0; i < 10; i++) {
    await new Promise((r) => setTimeout(r, 3000));
    const statusRes = await fetch(
      `https://graph.threads.net/v1.0/${creationId}?fields=status,error_message&access_token=${encodeURIComponent(accessToken)}`
    );
    if (!statusRes.ok) continue;
    const { status, error_message } = await statusRes.json();
    if (status === "FINISHED" || status === "PUBLISHED") break;
    if (status === "ERROR" || status === "EXPIRED") throw new Error(`threads container ${status}: ${error_message ?? ""}`);
  }

  const publishRes = await fetch(`https://graph.threads.net/v1.0/${userId}/threads_publish`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ creation_id: creationId, access_token: accessToken }),
  });
  if (!publishRes.ok) throw new Error(`threads publish failed: ${publishRes.status} ${await publishRes.text()}`);
  const { id } = await publishRes.json();
  let permalink = "";
  try {
    const p = await fetch(`https://graph.threads.net/v1.0/${id}?fields=permalink&access_token=${encodeURIComponent(accessToken)}`);
    if (p.ok) permalink = (await p.json()).permalink ?? "";
  } catch {
    // the post is live either way; the permalink is only for the calendar
  }
  return { id, permalink };
}
