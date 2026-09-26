#!/usr/bin/env node
// Rewrites the content of an already-live Blogger post from a local file.
// The file must be in posts/published/ and have a `blogger_url:` frontmatter
// field (added automatically by blogger-publish.mjs). Run via the "Update a
// published Blogger post" GitHub Action, with TARGET_FILE set to the file's
// path.
import { readFileSync, writeFileSync } from "node:fs";
import { basename } from "node:path";
import {
  parseFrontmatter,
  paragraphsToHtml,
  parseImagePrompts,
  aiImagePath,
  assetUrl,
  embedImages,
  getAccessToken,
  getBloggerPostByUrl,
  updateBloggerPost,
  deriveSearchDescription,
} from "./lib.mjs";

async function main() {
  const targetFile = process.env.TARGET_FILE;
  if (!targetFile) throw new Error("TARGET_FILE is not set");

  const raw = readFileSync(targetFile, "utf8");
  const { meta, body } = parseFrontmatter(raw);
  if (!meta.title) throw new Error(`${targetFile}: frontmatter is missing 'title'`);
  if (!meta.blogger_url) {
    throw new Error(
      `${targetFile}: frontmatter is missing 'blogger_url' — add it manually (the live post's URL) for posts published before this field existed.`
    );
  }
  const labels = meta.labels ? meta.labels.split(",").map((s) => s.trim()).filter(Boolean) : [];

  const accessToken = await getAccessToken();
  const existing = await getBloggerPostByUrl(meta.blogger_url, accessToken);

  const file = basename(targetFile);
  const images = parseImagePrompts(meta).map(({ alt }) => ({ url: assetUrl(aiImagePath(file, meta)), alt }));
  const html = embedImages(paragraphsToHtml(body), images).join("\n");
  const searchDescription = deriveSearchDescription(meta, body);

  const result = await updateBloggerPost({
    postId: existing.id,
    title: meta.title,
    html,
    labels,
    searchDescription,
    accessToken,
  });

  const updatedRaw = raw.replace(/^blogger_url:.*$/m, `blogger_url: ${result.url ?? meta.blogger_url}`);
  if (updatedRaw !== raw) writeFileSync(targetFile, updatedRaw);

  console.log(`Updated: ${meta.title} -> ${result.url ?? existing.url}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
