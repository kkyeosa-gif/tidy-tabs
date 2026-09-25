#!/usr/bin/env node
// Checks that published posts actually serve a <meta name="description">
// (see docs/search-description.md). Prints one line per post; exits 1 if any
// recent post is missing one.
import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { parseFrontmatter } from "./lib.mjs";

const DIR = "posts/published";
const LIMIT = Number(process.env.LIMIT || 10);
const files = readdirSync(DIR).filter((f) => f.endsWith(".md")).sort().slice(-LIMIT);
let missing = 0;
for (const f of files) {
  const { meta } = parseFrontmatter(readFileSync(join(DIR, f), "utf8"));
  if (!meta.blogger_url) continue;
  const html = await (await fetch(meta.blogger_url)).text();
  const m = html.match(/<meta content='([^']*)' name='description'\/>|<meta name='description' content='([^']*)'/);
  const desc = m ? (m[1] ?? m[2]) : "";
  if (!desc) missing += 1;
  console.log(`${desc ? "OK  " : "MISS"} ${f} ${desc.slice(0, 80)}`);
}
if (missing) {
  console.log(`${missing} post(s) without a meta description.`);
  process.exit(1);
}
