# Tidy Tabs style guide

Blog: Tidy Tabs (tidytabs.blogspot.com). Google Sheets and Excel templates and
how-tos for US small business owners and freelancers. Written in English.

## Voice
- Plain American English. Short paragraphs (1-3 sentences). Concrete verbs:
  click, select, type, drag, press. Not "leverage", "utilize", "navigate".
- Second person ("Select the column"). No filler intro, no "In this post we'll".
- Never repeat a point to hit a word count. 600-1100 words is normal; a
  short fix can be shorter.
- Never use: em dash, "moreover", "furthermore", "in essence", "it's worth
  noting", "dive into", "navigate", "landscape" (as a metaphor; the Landscape print setting is fine), "realm", "unlock", "elevate",
  "seamless", "delve", "game-changer", "Ultimate Guide", "Everything You Need
  to Know".

## Honesty rules (non-negotiable)
- Never write "I tested", "I tried", "in my experience" about something no one
  on the team actually did. Only claim a test that is recorded in the post's
  `tested_in:` field.
- If a behavior was checked in LibreOffice Calc but not Excel or Sheets, say
  exactly that. Excel/Sheets steps that come from Microsoft or Google help
  pages link to that page.
- Every screenshot is real and captioned with the app and version it came
  from. AI images (image_prompts) are illustrations only: never presented as a
  screenshot, never show UI text.
- Sample data is fictional and the post says so once.

## US reader details
- Money in dollars with two decimals: $18.00, $1,250.00.
- Dates as MM/DD/YYYY: 09/28/2026. Paper is US Letter (8.5 x 11 in).
- ZIP codes (02108, 07030), US states, Etsy / farmers market / 1099 client
  examples. No metric units, no A4.

## Post structure (in this order)
1. **Answer first.** The first paragraph is the direct answer in 1-2
   sentences, with the exact setting or formula. Example:
   "Keep ZIP codes intact by formatting the cells as Plain text before
   entering or pasting them. This preserves leading zeros in codes such as
   02108."
2. `> Works in:` callout: which apps/versions the steps cover and what was
   actually tested (mirror `tested_in`).
3. `## Steps` — numbered list, one action per step. Give menu paths as
   **File > Page Setup** in bold. Separate Excel and Google Sheets steps with
   `### In Excel` / `### In Google Sheets` when they differ.
4. `## Example` — a small table of US sample data showing before/after or the
   finished result, plus a real screenshot if one exists.
5. `## Troubleshooting` — 3-5 real failure cases as `### Symptom` headings,
   each with cause + fix in 1-3 sentences.
6. `## Template` — what's in the download, the download link
   (`[Download the .xlsx](templates/<file>.xlsx)`), and how to open it in
   Google Sheets (**File > Import > Upload**).

## Markdown the publisher understands
`## H2`, `### H3`, `- bullets`, `1. numbered`, `> callout`, `| pipe | tables |`,
inline `` `code` `` for formulas, `**bold**`, `[link](url)`,
and a standalone image line: `![alt text](images/<post-file-name>/shot.png) *caption*`.
Links/images without http are repo files and are published from the repo;
the post is held until every one of those files exists.

## Threads copy
Every post carries a `threads:` line (see .claude/agents/threads-writer.md).
It goes to Threads with the post link the moment the post publishes.
