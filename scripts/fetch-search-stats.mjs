#!/usr/bin/env node
// Pulls the last 28 days of Google Search Console data (top search queries
// and top pages) and writes a snapshot to tasks/search-stats.md. Runs
// weekly via the "Fetch search stats" GitHub Action so the analyst agent
// has real numbers instead of asking the user to paste them in by hand.
import { writeFileSync } from "node:fs";

const STATS_FILE = "tasks/search-stats.md";

async function getAccessToken() {
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.BLOGGER_CLIENT_ID,
      client_secret: process.env.BLOGGER_CLIENT_SECRET,
      refresh_token: process.env.GSC_REFRESH_TOKEN,
      grant_type: "refresh_token",
    }),
  });
  if (!res.ok) throw new Error(`token refresh failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.access_token;
}

async function queryStats({ accessToken, siteUrl, dimensions, startDate, endDate }) {
  const res = await fetch(
    `https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(siteUrl)}/searchAnalytics/query`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ startDate, endDate, dimensions, rowLimit: 25 }),
    }
  );
  if (!res.ok) throw new Error(`search console query failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.rows ?? [];
}

function toTable(rows, keyLabel) {
  if (!rows.length) return "_데이터 없음 (아직 노출/클릭이 없거나 반영 전)_\n";
  const header = `| ${keyLabel} | 클릭 | 노출 | CTR | 평균 순위 |\n| --- | --- | --- | --- | --- |\n`;
  const lines = rows.map((r) => {
    const key = r.keys[0];
    const ctr = (r.ctr * 100).toFixed(1) + "%";
    const pos = r.position.toFixed(1);
    return `| ${key} | ${r.clicks} | ${r.impressions} | ${ctr} | ${pos} |`;
  });
  return header + lines.join("\n") + "\n";
}

async function main() {
  const { BLOGGER_CLIENT_ID, BLOGGER_CLIENT_SECRET, GSC_REFRESH_TOKEN, GSC_SITE_URL } = process.env;
  if (!BLOGGER_CLIENT_ID || !BLOGGER_CLIENT_SECRET || !GSC_REFRESH_TOKEN || !GSC_SITE_URL) {
    console.log("Search Console secrets not fully set — skipping.");
    return;
  }

  const end = new Date();
  end.setDate(end.getDate() - 3); // GSC data lags a few days
  const start = new Date(end);
  start.setDate(start.getDate() - 28);
  const fmt = (d) => d.toISOString().slice(0, 10);

  const accessToken = await getAccessToken();
  const [queryRows, pageRows] = await Promise.all([
    queryStats({ accessToken, siteUrl: GSC_SITE_URL, dimensions: ["query"], startDate: fmt(start), endDate: fmt(end) }),
    queryStats({ accessToken, siteUrl: GSC_SITE_URL, dimensions: ["page"], startDate: fmt(start), endDate: fmt(end) }),
  ]);

  const totalClicks = queryRows.reduce((sum, r) => sum + r.clicks, 0);
  const totalImpressions = queryRows.reduce((sum, r) => sum + r.impressions, 0);

  const content = `# Search stats (Google Search Console)

analyst 에이전트가 여기를 읽는다. 매주 자동 갱신되는 스냅샷이라, 이전 기록은
git 히스토리(\`git log -p -- tasks/search-stats.md\`)에서 볼 수 있다.

기간: ${fmt(start)} ~ ${fmt(end)} (최근 28일)
합계: 클릭 ${totalClicks}회, 노출 ${totalImpressions}회

## 어떤 검색어로 들어오는지

${toTable(queryRows, "검색어")}

## 어떤 글이 잘 나오는지

${toTable(pageRows, "페이지")}
`;

  writeFileSync(STATS_FILE, content);
  console.log(`Wrote ${STATS_FILE}: ${queryRows.length} queries, ${pageRows.length} pages.`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
