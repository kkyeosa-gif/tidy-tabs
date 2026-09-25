#!/usr/bin/env node
// Generates the AI illustrations for every queued post (posts/drafts/ and
// posts/ready/) that doesn't have them yet: one image per `image_prompts`
// scene, saved to images/<post-slug>/ai-<n>.jpg. Existing files are never
// regenerated, so each image is paid for once and stays the same after the
// team has reviewed it. Delete a file to re-roll it.
import { readFileSync, readdirSync, existsSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { parseFrontmatter, parseImagePrompts, aiImagePath, generateAiImage } from "./lib.mjs";

const DIRS = ["posts/drafts", "posts/ready"];

async function main() {
  const todo = [];
  for (const dir of DIRS) {
    if (!existsSync(dir)) continue;
    for (const file of readdirSync(dir).filter((f) => f.endsWith(".md")).sort()) {
      const { meta } = parseFrontmatter(readFileSync(join(dir, file), "utf8"));
      for (const { n, prompt } of parseImagePrompts(meta)) {
        const out = aiImagePath(file, n);
        if (!existsSync(out)) todo.push({ file, out, prompt });
      }
    }
  }
  if (!todo.length) {
    console.log("Every queued post already has its AI images.");
    return;
  }
  if (!process.env.OPENAI_API_KEY) throw new Error(`OPENAI_API_KEY is not set (${todo.length} image(s) missing)`);

  let failed = 0;
  for (const { file, out, prompt } of todo) {
    try {
      const jpg = await generateAiImage(prompt);
      mkdirSync(dirname(out), { recursive: true });
      writeFileSync(out, jpg);
      console.log(`Generated ${out} (${Math.round(jpg.length / 1024)} KB) for ${file}`);
    } catch (err) {
      failed += 1;
      console.warn(`${file}: ${err.message}`);
    }
  }
  // Images that did generate still get committed by the workflow; a post
  // whose image failed is simply held by the publisher until a re-run.
  if (failed) console.warn(`${failed} image(s) failed; re-run to retry.`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
