#!/usr/bin/env node
// One-time, LOCAL-ONLY helper: mints a Google API refresh token via the
// OAuth "installed app" loopback flow. Never run this in CI — it opens a
// local server and needs an interactive Google login in your own browser.
//
// Usage (Blogger, the default):
//   BLOGGER_CLIENT_ID=... BLOGGER_CLIENT_SECRET=... node scripts/get-refresh-token.mjs
//
// Usage (any other API, e.g. Search Console — reuse the same client ID/secret,
// just enable the API in the same Google Cloud project first):
//   BLOGGER_CLIENT_ID=... BLOGGER_CLIENT_SECRET=... \
//   SCOPES="https://www.googleapis.com/auth/webmasters.readonly" \
//   node scripts/get-refresh-token.mjs
//
// It prints a URL — open it, log in with the Google account that owns the
// blog, approve access, and the script prints the refresh token to save as
// a GitHub secret.
import { createServer } from "node:http";

const CLIENT_ID = process.env.BLOGGER_CLIENT_ID;
const CLIENT_SECRET = process.env.BLOGGER_CLIENT_SECRET;
const SCOPES = process.env.SCOPES || "https://www.googleapis.com/auth/blogger";
const PORT = 53682;
const REDIRECT_URI = `http://127.0.0.1:${PORT}/callback`;

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error("Set BLOGGER_CLIENT_ID and BLOGGER_CLIENT_SECRET first.");
  process.exit(1);
}

const authUrl = new URL("https://accounts.google.com/o/oauth2/v2/auth");
authUrl.searchParams.set("client_id", CLIENT_ID);
authUrl.searchParams.set("redirect_uri", REDIRECT_URI);
authUrl.searchParams.set("response_type", "code");
authUrl.searchParams.set("scope", SCOPES);
authUrl.searchParams.set("access_type", "offline");
authUrl.searchParams.set("prompt", "consent");

console.log("\n1) Open this URL in your browser and approve access:\n");
console.log(authUrl.toString());
console.log(`\n2) Waiting for the redirect on ${REDIRECT_URI} ...\n`);

const server = createServer(async (req, res) => {
  const url = new URL(req.url, REDIRECT_URI);
  if (url.pathname !== "/callback") {
    res.writeHead(404).end();
    return;
  }
  const code = url.searchParams.get("code");
  res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
  res.end("<h1>인증 완료</h1><p>이 창은 닫아도 됩니다. 터미널을 확인하세요.</p>");
  server.close();

  const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      code,
      grant_type: "authorization_code",
      redirect_uri: REDIRECT_URI,
    }),
  });
  const data = await tokenRes.json();
  if (!tokenRes.ok) {
    console.error("Token exchange failed:", data);
    process.exit(1);
  }
  console.log("\nSave this refresh token as the matching GitHub secret");
  console.log("(BLOGGER_REFRESH_TOKEN for the default scope, GSC_REFRESH_TOKEN for Search Console, etc.):\n");
  console.log(data.refresh_token);
  console.log();
});

server.listen(PORT);
