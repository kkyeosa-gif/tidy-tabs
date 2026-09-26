---
title: How to Keep Leading Zeros in Excel ZIP Codes
labels: excel-basics, data-entry
tested_in: LibreOffice Calc 24.2 (Linux), 09/2026. Excel and Google Sheets steps from Microsoft/Google help pages.
image_prompts: A stack of addressed US mail envelopes beside a laptop with a spreadsheet open, labels and screen out of focus
image_alt: Addressed envelopes stacked beside a laptop with a spreadsheet open
threads: Excel ate the leading 0 in your ZIP codes again?\nFormat the column as Text before you paste, not after.
search_description: Excel dropping the 0 from ZIP codes like 02108? Format the column as Text first, or rebuild it with =TEXT(A2,"00000").
---
Answer first, in 1-2 sentences, with the exact setting or formula.

> Works in: Excel for Microsoft 365 (Windows/Mac), Google Sheets. Tested in: see tested_in.

![ZIP codes stored as numbers next to the same codes repaired with TEXT](images/<post-file>/zip-codes-leading-zeros-fixed.png) *Where the image came from (app and version).*

## Steps
1. One action per step. Menu paths in **bold**.

## Example
| Before | After |
| --- | --- |
| 2108 | 02108 |

## Troubleshooting
### The zeros disappear again after I paste
Cause and fix in 1-3 sentences.

## Template
[Download the .xlsx](templates/file-name.xlsx)
