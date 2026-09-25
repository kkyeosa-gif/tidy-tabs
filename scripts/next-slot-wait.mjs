#!/usr/bin/env node
// Prints how many seconds the publish clock should sleep: until 3 minutes
// after the next publish slot (America/Los_Angeles), capped so one job stays
// under GitHub's 6-hour limit. Also prints whether that wake-up is a slot.
// Output (GITHUB_OUTPUT format): seconds=<n> and is_slot=<true|false>.
const SLOT_HOURS = [5, 8, 10, 13, 17]; // keep in sync with blogger-publish.mjs
const CAP = 5 * 3600 + 30 * 60;

function laParts(d) {
  const p = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Los_Angeles", year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23",
  }).formatToParts(d);
  const g = (t) => Number(p.find((x) => x.type === t).value);
  return { h: g("hour") % 24, m: g("minute"), s: g("second") };
}

const now = new Date();
const { h, m, s } = laParts(now);
const secOfDay = h * 3600 + m * 60 + s;
const targets = [...SLOT_HOURS, ...SLOT_HOURS.map((x) => x + 24)].map((x) => x * 3600 + 180);
const next = targets.find((t) => t > secOfDay + 30);
let wait = next - secOfDay;
const isSlot = wait <= CAP;
if (!isSlot) wait = CAP;
console.log(`seconds=${wait}`);
console.log(`is_slot=${isSlot}`);
