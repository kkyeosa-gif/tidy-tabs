---
threads_url: https://www.threads.com/@tin_ylab/post/DgAnYu9FKeY
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-net-promoter-score.html
title: How to Calculate a Net Promoter Score From a 0 to 10 Customer Survey in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 22 workbooks and 22,968 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Get NPS from 0 to 10 survey scores: percent of 9s and 10s minus percent of 0 to 6, with COUNTIF on a cleaned score column, by channel too. Free template.
image_prompts: A small wooden drop box with a slot in the lid on a craft booth checkout table, a short stack of blank unfilled paper feedback cards beside it, and a cup of short pencils, an empty outdoor market in late afternoon light, no people, realistic photo, no screens, no readable text, no logos
image_alt: Wooden slotted drop box beside blank feedback cards and a cup of pencils
threads: COUNTA counts text, and your NPS pays for it.\nIn a made-up set of 20 scores, an 11 and a text "9" moved the score from 30 to 31.82.\nCount a helper column of whole numbers 0 to 10 with COUNT: =(COUNTIF(scores,">=9")-COUNTIF(scores,"<=6"))/COUNT(scores)*100
---
Subtract the percent of detractors (0 to 6) from the percent of promoters (9 and 10) with `=(COUNTIF(scores,">=9")-COUNTIF(scores,"<=6"))/COUNT(scores)*100`, run on a helper column that keeps only whole numbers from 0 to 10. In the made-up set of 20 scores, 11 promoters (55%) minus 5 detractors (25%) gives an NPS of 30, while the same formula with COUNTA on the raw column gives 30.43 once an 11, an 8.5 and a text "9" sneak in; the download marks those rows Check.

Passives (7 and 8) count in the total but in neither percent. The result is a whole number from -100 to 100. Net Promoter Score and NPS are trademarks of their owners (Bain & Company, Satmetrix/NICE and Fred Reichheld); this post uses the terms only to describe the calculation and is not connected with them.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (22 workbooks, 22,968 cell comparisons, 0 mismatches), plus the edge cases below. The formulas use plain functions (COUNT, COUNTIF, COUNTIFS, ISNUMBER, ROUND) that also exist in Excel and Google Sheets, but neither app was opened for this post.

![Survey summary by channel with promoter, passive and detractor counts, percents and score](images/2026-11-17-team-04-net-promoter-score-from-customer-survey-0-10-countifs-excel-google-sheets/nps-score-by-channel-template.png) *Twenty fictional survey scores: 11 promoters (55.0%) minus 5 detractors (25.0%) gives 30 overall, with Etsy at 50, Market at 14 and Email at 20 (the unrounded column and check rows are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Put one survey answer per row on a **Responses** tab, starting in row 5.

1. Type or paste the date in **A5** (11/02/2026), the channel in **B5** (Etsy, Market or Email) and the 0 to 10 score in **C5**. Paste scores as numbers, not text.
2. In **E5**, add a helper that keeps only valid scores: `=IF(AND(ISNUMBER(C5),C5>=0,C5<=10),IF(C5=ROUND(C5,0),C5,""),"")`. A text "9", an 8.5, an 11 or a -1 comes out blank, so it is not counted anywhere.
3. In **F5**, label each row: `=IF(C5="","",IF(ISNUMBER(E5),IF(E5>=9,"Promoter",IF(E5>=7,"Passive","Detractor")),"Check"))`. A blank score stays blank; a bad one shows Check.
4. Fill **E5:F5** down to row 504.
5. On a **Summary** tab, count the valid scores with `=COUNT(Responses!$E$5:$E$504)`. Use COUNT on the helper, not COUNTA on column C: COUNTA also counts text.
6. Count promoters with `=COUNTIF(Responses!$E$5:$E$504,">=9")`, passives with `=COUNTIFS(Responses!$E$5:$E$504,">=7",Responses!$E$5:$E$504,"<=8")` and detractors with `=COUNTIF(Responses!$E$5:$E$504,"<=6")`.
7. Work out the score with `=IF(N(B5)=0,"",ROUND((C5-E5)/B5*100,0))`, where B5 is the count, C5 the promoters and E5 the detractors. The IF keeps it blank instead of showing #DIV/0! before any scores are in.
8. For each channel, add the channel as a second condition: `=COUNTIFS(Responses!$B$5:$B$504,$A6,Responses!$E$5:$E$504,">=9")`, with the channel name in **A6**. Repeat with ">=0" for the count, and with the passive and detractor bands, then use the same score formula on that row.

Banding scores with COUNTIFS works the same way as in [counting orders by dollar amount range](https://tidytabs.blogspot.com/2026/10/how-to-count-orders-by-dollar-amount.html).

Help pages: Microsoft's [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34), [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [COUNT function](https://support.microsoft.com/en-us/office/count-function-a59cd7fc-b623-4d93-87a4-d23bf411294c).

## Example

The responses are fictional: 20 scores dated 11/02/2026 to 11/11/2026 from three channels. The first few rows look like this.

| Date | Channel | Score 0-10 | Comment | Valid score | Group |
|---|---|---|---|---|---|
| 11/02/2026 | Etsy | 10 | Fast shipping | 10 | Promoter |
| 11/02/2026 | Etsy | 9 | | 9 | Promoter |
| 11/03/2026 | Market | 9 | | 9 | Promoter |
| 11/03/2026 | Etsy | 8 | | 8 | Passive |
| 11/04/2026 | Etsy | 6 | Box arrived dented | 6 | Detractor |

The Summary tab turns all 20 into one table:

| Channel | Counted | Promoters | Passives | Detractors | % promoters | % detractors | NPS |
|---|---|---|---|---|---|---|---|
| All responses | 20 | 11 | 4 | 5 | 55.0% | 25.0% | 30 |
| Etsy | 8 | 5 | 2 | 1 | 62.5% | 12.5% | 50 |
| Market | 7 | 3 | 2 | 2 | 42.9% | 28.6% | 14 |
| Email | 5 | 3 | 0 | 2 | 60.0% | 40.0% | 20 |

The Market score is 14.29 before rounding; the "NPS (not rounded)" column shows it. Adding a text "9", a blank, an 11, an 8.5, a -1 and a text " 10" to the list still gave 20 counted and an NPS of 30, with 5 rows marked Check and the blank row left blank.

## Troubleshooting

### The score is a little off
The count came from COUNTA on the raw column, which counts text and out-of-range entries. On the sample, adding an 11 and a text "9" moved the score from 30 to 31.82, and adding an 8.5 as well gave 30.43. Count with COUNT on the validated helper column instead.

### A row shows Check
The score is not a whole number from 0 to 10: a number stored as text, a decimal such as 8.5, an 11 or a -1. Retype it as a plain number. The row stays out of every count until you do.

### The channel scores do not add up to the total
A channel name has an extra space ("Etsy ") or is not one of the names in **Summary A6:A9**. In LibreOffice, "etsy" and "ETSY" still counted as Etsy, but "Etsy " did not. **B12** shows how many counted scores have no listed channel; [TRIM](https://tidytabs.blogspot.com/2026/10/how-to-clean-up-customer-list-with-trim.html) removes the stray spaces.

### NPS is blank
There are no valid scores yet, so the count is 0. The template leaves the cell blank instead of showing #DIV/0!. Check that the scores are in column C as numbers.

## Template

[Download the .xlsx](templates/tidy-tabs-nps-survey-score.xlsx)

The file has three tabs. **Responses** holds rows 5 to 504: Date, Channel (a dropdown from Summary A6:A9), Score 0-10 and Comment (inputs in **A:D**), then Valid score and Group (formulas in **E:F**). **Summary** has one row for all responses and rows 6 to 9 for channels (type your own names in **A6:A9**; the fourth row is blank for a channel you add), with Responses counted, Promoters, Passives, Detractors, % promoters, % detractors, NPS and NPS (not rounded). Below it, **B11:B14** show rows showing Check, counted scores with no listed channel, entries typed below row 504 and the CHECK. **How to use** explains each column and carries the trademark note. The data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. Promoter cells turn green, Check rows turn red, and so does the CHECK cell when it is not OK. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
