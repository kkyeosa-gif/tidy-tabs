#!/usr/bin/env node
// Posts to Threads for published blog posts that don't have a threads_url
// yet (Threads step failed, or the token wasn't set when they went live).
// At most MAX_PER_RUN per run, oldest first: a brand-new Threads account
// that fires many API calls at once can get "API access blocked" (see
// half-handy docs/threads-setup.md), so this catches up slowly.
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { parseFrontmatter } from "./lib.mjs";
import { postToThreads } from "./blogger-publish.mjs";

const PUBLISHED_DIR = "posts/published";
const MAX_PER_RUN = Number(process.env.MAX_PER_RUN || 2);

const pending = readdirSync(PUBLISHED_DIR)
  .filter((f) => f.endsWith(".md"))
  .sort()
  .map((file) => ({ file, ...parseFrontmatter(readFileSync(join(PUBLISHED_DIR, file), "utf8")) }))
  .filter(({ meta }) => meta.blogger_url && !meta.threads_url);

console.log(`${pending.length} published post(s) without a Threads post.`);
let done = 0;
for (const { file, meta, body } of pending.slice(0, MAX_PER_RUN)) {
  const post = await postToThreads({ file, meta, body, url: meta.blogger_url, backfill: true });
  if (!post) process.exitCode = 1;
  else done += 1;
}
console.log(`Backfilled ${done}.`);
