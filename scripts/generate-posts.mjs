#!/usr/bin/env node
// Writes first drafts into posts/drafts/ (never posts/ready/). Manual-only
// via the "Draft posts" GitHub Action: every Tidy Tabs post promises a
// tested template and real screenshots, which a model can't produce, so a
// person/the content team has to finish each draft before it's queued.
// Two Claude calls per post: draft, then a technical fact-check pass.
import { readFileSync, readdirSync, writeFileSync, existsSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { parseFrontmatter, factCheckAndFix } from "./lib.mjs";

const DRAFTS_DIR = "posts/drafts";
const READY_DIR = "posts/ready";
const PUBLISHED_DIR = "posts/published";
const TOPICS_FILE = "tasks/topics-seed.md";
const STYLE_FILE = "docs/style-guide.md";
const COUNT = Math.max(1, Math.min(5, Number(process.env.DRAFT_COUNT || 1)));

function listPostFiles(dir) {
  if (!existsSync(dir)) return [];
  return readdirSync(dir).filter((f) => f.endsWith(".md") && f !== "TEMPLATE.md");
}

function usedTitlesAndLabels() {
  const files = [
    ...listPostFiles(DRAFTS_DIR).map((f) => join(DRAFTS_DIR, f)),
    ...listPostFiles(READY_DIR).map((f) => join(READY_DIR, f)),
    ...listPostFiles(PUBLISHED_DIR).map((f) => join(PUBLISHED_DIR, f)),
  ];
  const titles = [];
  const labels = new Set();
  for (const f of files) {
    try {
      const { meta } = parseFrontmatter(readFileSync(f, "utf8"));
      if (meta.title) titles.push(meta.title);
      (meta.labels ?? "")
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean)
        .forEach((l) => labels.add(l));
    } catch {
      // skip unparseable files rather than failing the whole run
    }
  }
  return { titles, labels: [...labels] };
}

function slugify(title) {
  return title
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40);
}

async function draftPost({ usedTitles, usedLabels, topicsSeed, styleGuide }) {
  const prompt = `You are drafting one post for Tidy Tabs, an English-language blog of Google Sheets and Microsoft Excel templates and how-tos for US small business owners and freelancers.

Already-published or queued post titles (do not repeat these topics):
${usedTitles.map((t) => `- ${t}`).join("\n") || "(none yet)"}

Topic ideas (pick one not covered above, or a similarly concrete spreadsheet task a US small business owner would search for):
${topicsSeed}

Follow this style guide exactly:
${styleGuide}

Extra rules for a machine draft:
- Never write "I tested", "I tried", or any claim of hands-on experience. You haven't tested anything. Write instructions in the second person.
- Where a real screenshot belongs, write a line exactly like: SCREENSHOT NEEDED: <what the screen should show, which app and version>
- Where the downloadable template goes, write: TEMPLATE NEEDED: <file name and what columns/formulas it contains>
- Write labels: 1-3 short tags, avoiding these existing ones unless the topic genuinely overlaps: ${usedLabels.join(", ") || "(none yet)"}.

Output ONLY the raw file content in exactly this format, nothing else — no code fences, no explanation:

---
title: <title>
labels: <label1>, <label2>
tested_in: NOT TESTED YET
image_prompts: <illustration scene 1> | <illustration scene 2>
image_alt: <alt text 1> | <alt text 2>
search_description: <description>
---
<body>`;

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
      messages: [{ role: "user", content: prompt }],
    }),
  });
  if (!res.ok) throw new Error(`post generation failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  let text = (data.content?.find((b) => b.type === "text")?.text ?? "").trim();
  if (text.startsWith("```")) {
    text = text.replace(/^```[a-z]*\n/, "").replace(/```\s*$/, "").trim();
  }
  if (!text.startsWith("---")) throw new Error(`unexpected generation response: ${text.slice(0, 200)}`);
  return text;
}

async function main() {
  if (!process.env.ANTHROPIC_API_KEY) {
    console.log("ANTHROPIC_API_KEY not set — skipping draft generation.");
    return;
  }
  mkdirSync(DRAFTS_DIR, { recursive: true });
  const topicsSeed = existsSync(TOPICS_FILE) ? readFileSync(TOPICS_FILE, "utf8") : "";
  const styleGuide = existsSync(STYLE_FILE) ? readFileSync(STYLE_FILE, "utf8") : "";
  const today = new Date().toISOString().slice(0, 10);

  for (let i = 0; i < COUNT; i++) {
    try {
      const { titles: usedTitles, labels: usedLabels } = usedTitlesAndLabels();
      const draft = await draftPost({ usedTitles, usedLabels, topicsSeed, styleGuide });
      let fixed = draft;
      try {
        fixed = await factCheckAndFix(draft);
      } catch (err) {
        console.warn(`Fact-check pass failed, keeping the unchecked draft (it's only a draft): ${err.message}`);
      }
      const titleMatch = fixed.match(/^title:\s*(.+)$/m);
      const title = titleMatch ? titleMatch[1].trim() : `post-${Date.now()}`;
      const filename = `${today}-draft-${String(i + 1).padStart(2, "0")}-${slugify(title)}.md`;
      writeFileSync(join(DRAFTS_DIR, filename), fixed + "\n");
      console.log(`Drafted: ${filename} (${title})`);
    } catch (err) {
      console.warn(`Skipping draft ${i + 1}/${COUNT}: ${err.message}`);
    }
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
