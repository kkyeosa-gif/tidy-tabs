#!/usr/bin/env node
// Publishes the next ready post to Blogger once per SLOT_HOURS slot
// (America/Los_Angeles). Meant to run from a GitHub Actions cron that fires
// several times an hour; the first run at or after a slot starts publishes,
// and a per-day/per-slot marker file makes later runs (cron retries, manual
// dispatch) no-op so nothing double-posts.
import { readFileSync, readdirSync, existsSync, mkdirSync, unlinkSync, writeFileSync, appendFileSync } from "node:fs";
import { join } from "node:path";
import {
  parseFrontmatter,
  paragraphsToHtml,
  parseImagePrompts,
  aiImagePath,
  assetUrl,
  localAssetRefs,
  embedImages,
  getAccessToken,
  publishToBlogger,
  deriveSearchDescription,
  translateToKorean,
  createNotionPage,
} from "./lib.mjs";

// One post a day, 6am Pacific (9am Eastern). A brand-new blogspot that posts
// five AI-assisted articles a day looks like scaled content to Google (see
// half-handy's 2026-09-24 analysis), and every post here needs a tested
// template and real screenshots, so volume isn't the goal.
const SLOT_HOURS = [6];
const READY_DIR = "posts/ready";
const PUBLISHED_DIR = "posts/published";
const MARKER_DIR = "tasks/.publish-markers";
const CALENDAR_FILE = "tasks/calendar.md";

function laNow() {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Los_Angeles",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    hourCycle: "h23",
  }).formatToParts(new Date());
  const get = (t) => parts.find((p) => p.type === t).value;
  return { date: `${get("year")}-${get("month")}-${get("day")}`, hour: Number(get("hour")) % 24 };
}

// Every file the post needs (AI images, screenshots, templates) must already
// be committed, or readers get broken images and dead download links.
function missingAssets(file, meta, body) {
  const needed = [
    ...parseImagePrompts(meta).map(({ n }) => aiImagePath(file, n)),
    ...localAssetRefs(body),
  ];
  return needed.filter((p) => !existsSync(p));
}

// Oldest ready post whose assets are all in place. A post still waiting on
// a screenshot is skipped (and logged) instead of blocking the whole queue.
function nextPublishable() {
  if (!existsSync(READY_DIR)) return null;
  const files = readdirSync(READY_DIR).filter((f) => f.endsWith(".md")).sort();
  for (const file of files) {
    const raw = readFileSync(join(READY_DIR, file), "utf8");
    const { meta, body } = parseFrontmatter(raw);
    const missing = missingAssets(file, meta, body);
    if (missing.length) {
      console.warn(`Holding ${file}: missing ${missing.join(", ")}`);
      continue;
    }
    return { file, raw, meta, body };
  }
  return null;
}

async function main() {
  const { date, hour } = laNow();
  const force = process.env.FORCE_PUBLISH === "true";
  const slot = [...SLOT_HOURS].reverse().find((h) => h <= hour);
  if (slot === undefined && !force) {
    console.log(`LA time is ${date} ${hour}:00 — before the publish slot (${SLOT_HOURS.join(", ")}). Skipping.`);
    return;
  }

  mkdirSync(MARKER_DIR, { recursive: true });
  const marker = join(MARKER_DIR, `${date}-${force ? hour : slot}.done`);
  if (existsSync(marker) && !force) {
    console.log(`Already published for the ${date} ${slot}:00 slot. Skipping.`);
    return;
  }

  const next = nextPublishable();
  if (!next) {
    console.log("No publishable posts in posts/ready/. Nothing to publish.");
    return;
  }
  const { file, raw, meta, body } = next;
  if (!meta.title) throw new Error(`${file}: frontmatter is missing 'title'`);
  const labels = meta.labels ? meta.labels.split(",").map((s) => s.trim()).filter(Boolean) : [];

  const images = parseImagePrompts(meta).map(({ n, alt }) => ({ url: assetUrl(aiImagePath(file, n)), alt }));
  const html = embedImages(paragraphsToHtml(body), images).join("\n");
  const searchDescription = deriveSearchDescription(meta, body);

  const accessToken = await getAccessToken();
  const result = await publishToBlogger({ title: meta.title, html, labels, searchDescription, accessToken });

  mkdirSync(PUBLISHED_DIR, { recursive: true });
  // Record the live URL in the moved file's frontmatter so a future edit
  // (scripts/blogger-update.mjs) can find the post without re-searching.
  const publishedRaw = raw.replace(/^---\n/, `---\nblogger_url: ${result.url ?? ""}\n`);
  writeFileSync(join(PUBLISHED_DIR, file), publishedRaw);
  unlinkSync(join(READY_DIR, file));
  writeFileSync(marker, `${new Date().toISOString()} -> ${result.url ?? result.id}\n`);
  appendFileSync(
    CALENDAR_FILE,
    `\n- ${date} ${String(hour).padStart(2, "0")}:00 PT — published "${meta.title}" (${file}, ${images.length} AI image${images.length === 1 ? "" : "s"}) -> ${result.url ?? result.id}`
  );

  console.log(`Published: ${meta.title} -> ${result.url ?? result.id}`);

  // Korean translation goes under the same Notion parent page as the other
  // blogs, so the title gets a "-tidy-tabs" suffix to tell them apart.
  // Best-effort: never fail the job after the post is live (that would skip
  // the commit step and leave the repo out of sync with Blogger).
  if (!process.env.ANTHROPIC_API_KEY || !process.env.NOTION_API_KEY || !process.env.NOTION_PARENT_PAGE_ID) {
    console.warn("Skipping Korean Notion page: ANTHROPIC_API_KEY / NOTION_API_KEY / NOTION_PARENT_PAGE_ID not fully set.");
  } else {
    try {
      const translated = await translateToKorean(meta.title, body);
      const notionPage = await createNotionPage({
        title: `${translated.title}-tidy-tabs`,
        body: translated.body,
        sourceTitle: meta.title,
        sourceUrl: result.url ?? "",
      });
      console.log(`Notion page created: ${notionPage.url ?? notionPage.id}`);
    } catch (err) {
      console.warn(`Notion translation step failed (blog post already published, continuing): ${err.message}`);
    }
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
