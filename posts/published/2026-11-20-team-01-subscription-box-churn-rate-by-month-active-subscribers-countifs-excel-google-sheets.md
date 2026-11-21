---
threads_url: https://www.threads.com/@tin_ylab/post/DgBJtjTlKXd
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-calculate-monthly-churn-rate-for.html
title: How to Calculate Monthly Churn Rate for a Subscription Box in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 37 workbooks and 27,898 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Count active subscribers on the 1st with COUNTIFS (blank cancel date or cancel on/after the 1st), then divide that group's cancels by it. Free template.
image_prompts: A home studio packing table with neat rows of identical kraft paper subscription mailer boxes filled with tissue paper and small wrapped soap bars, one box pulled out of the row sitting empty with its lid off, and a roll of blank white shipping labels beside it, soft window light, realistic photo, no screens, no readable text, no logos
image_alt: Rows of packed kraft subscription boxes with one empty box set aside
threads: July started with 4 subscribers and 1 of them cancelled: 25.00% churn.\nDivide all 2 of July's cancels by 4 and you get 50.00%, because one made-up subscriber joined and left inside July.\nCount the start group with two COUNTIFS: blank cancel date, plus a cancel on or after the 1st.
---
Count the subscribers active on the 1st of each month with `=COUNTIFS(start,"<"&m,cancel,"")+COUNTIFS(start,"<"&m,cancel,">="&m)`, which adds the ones with no cancel date to the ones who cancel on or after that day, then divide the people from that group who cancel during the month by it. In the made-up list of 14 subscribers, July starts with 4, one of them cancels, and the churn rate is 25.00%, while dividing all of July's 2 cancels by 4 gives 50.00% because one of those subscribers joined and left inside July.

Churn has more than one definition. This sheet uses cancels from the start group divided by the start count, and shows the all-cancels version only as a comparison. It counts subscribers, not revenue.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (37 workbooks, 27,898 cell comparisons, 0 mismatches), plus the edge cases below. The functions also exist in Excel and Google Sheets, but neither app was opened for this post. Confirm there that the `""` criterion counts an empty cancel cell and that AVERAGEIFS skips blank churn cells.

![Six months of subscriber counts with start-group churn beside an all-cancels comparison rate](images/2026-11-20-team-01-subscription-box-churn-rate-by-month-active-subscribers-countifs-excel-google-sheets/subscription-churn-rate-by-month-template.png) *Fictional subscribers from June to November 2026: in July 4 were active at the start, 2 cancelled but only 1 of them came from the start group, so churn is 25.00%, not the 50.00% you get by dividing all cancels by the start count (the End = next start? and blank-only count columns, the summary, the average-through row and the months from December on are hidden here, not in the file). PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets.

1. On the **Subscribers** tab, type one subscriber per row starting in row 5: name or ID in **A**, plan in **B**, start date in **C** and cancel date in **D** (MM/DD/YYYY).
2. Leave the cancel date empty while the subscriber is active. The cancel date is the day it stopped. A pause and restart is a new row.
3. On the **Monthly** tab, type any date in your first month in **B1**. The sheet moves it to the 1st and lists 24 months in **A5:A28**.
4. In **B5**, count who was active on the 1st: `=COUNTIFS(Subscribers!$C$5:$C$504,"<"&A5,Subscribers!$D$5:$D$504,"")+COUNTIFS(Subscribers!$C$5:$C$504,"<"&A5,Subscribers!$D$5:$D$504,">="&A5)`. The second part catches people who cancel on or after the 1st, because they were still subscribers that morning.
5. **C5** and **D5** count start dates and cancel dates from this month's 1st up to next month's 1st (**A6**).
6. In **E5**, count only the cancels that came from the start group: `=COUNTIFS(Subscribers!$C$5:$C$504,"<"&A5,Subscribers!$D$5:$D$504,">="&A5,Subscribers!$D$5:$D$504,"<"&A6)`.
7. In **G5**, divide: `=IF(N(B5)=0,"",E5/B5)`. A month with 0 at the start stays blank.
8. **I5** shows OK when Active at end (**F5**, start + new - cancelled) equals the next month's start.
9. Type the last full month you want averaged in **B2**. **M6** then averages the churn rates from **B1** to that month, skipping months with a start count of 0.

Microsoft documents the criteria rules on its [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [AVERAGEIFS function](https://support.microsoft.com/en-us/office/averageifs-function-48910c45-1fc0-4389-a028-f7c5c3001690) pages, and Google on its [COUNTIFS](https://support.google.com/docs/answer/3256550) page. For each subscriber's next renewal date, see [tracking subscription renewals](https://tidytabs.blogspot.com/2026/10/how-to-track-subscription-renewals-in.html).

## Example

The 14 subscribers, Ava to Nora, are fictional. With **B1** set to 06/01/2026, the Monthly tab shows:

| Month start | Active at start | New | Cancelled | From start group | Active at end | Churn rate | All cancels / start | Blank-only count |
|---|---|---|---|---|---|---|---|---|
| 06/01/2026 | 0 | 4 | 0 | 0 | 4 | (blank) | (blank) | 0 |
| 07/01/2026 | 4 | 3 | 2 | 1 | 5 | 25.00% | 50.00% | 2 |
| 08/01/2026 | 5 | 3 | 1 | 1 | 7 | 20.00% | 20.00% | 3 |
| 09/01/2026 | 7 | 2 | 2 | 1 | 7 | 14.29% | 28.57% | 5 |
| 10/01/2026 | 7 | 2 | 1 | 1 | 8 | 14.29% | 14.29% | 6 |
| 11/01/2026 | 8 | 0 | 1 | 1 | 7 | 12.50% | 12.50% | 7 |

Every End = next start? cell shows OK. Nora's cancel date is 07/01/2026, so she counts as active at the start of July and cancelled in July. Dev joined on 07/01/2026 and left on 07/25/2026, so he was never in July's start count of 4. Counting his cancel turns 25.00% into 50.00%.

With **B2** at 10/01/2026, the summary in **L4:M12** shows 4 months in the average, an average monthly churn rate of 18.39%, 14 new subscribers, 6 cancellations and CHECK OK.

## Troubleshooting

### Active at start is too low
Only blank cancel dates were counted. `COUNTIFS(start,"<"&m,cancel,"")` alone gave July 2 instead of 4 (column **J**), because Ben cancelled on 08/10/2026 and Nora on 07/01/2026, and both were active on 07/01. Add the second COUNTIFS for cancel dates on or after the 1st.

### Churn looks too high in a month with sign-ups
All of the month's cancels were divided by the start count. That gave 50.00% for July and 28.57% for September (column **H**), because Dev and Jon joined and left inside those months. Use only cancels from the start group (column **E**).

### End = next start? shows Check
A cancel date is typed as text, the cancel cell holds a space, a cancel date has no start date, or a cancel date is before the start. In a copy with a 09/10/2026 cancel and no start date added, September showed 3 cancelled and an end of 6 against October's start of 7, and the row's Status showed Check: cancel, no start. Fix the date, or clear a space with Delete.

### The average churn is lower than you expect
**B2** is later than your data, so empty months get averaged: from December 2026 on, the sample shows 7 active and 0.00% every month. With **B2** blank, the sheet stops at the month of the latest date typed, 11/01/2026 here, and gave 17.21% over 5 months instead of 18.39% through 10/01/2026. Set **B2** to the last month you want counted.

## Template

[Download the .xlsx](templates/tidy-tabs-subscription-churn-rate.xlsx)

The file has three tabs. **Subscribers** has rows 5 to 504 with Subscriber or ID, Plan, Start date and Cancel date (inputs in **A:D**) and a Status in **E**: Active, Cancelled, or a Check that names the problem, such as Check: cancel before start. **G4:H8** counts subscribers, active, cancelled, Check rows and entries below row 504.

**Monthly** has First month in **B1**, Average churn through in **B2**, 24 months in rows 5 to 28 with the example's columns plus End = next start?, and the summary in **L4:M12**. **How to use** explains each formula. Both data tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas. A Status starting with Check, an End = next start? Check and a CHECK that is not OK turn red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
