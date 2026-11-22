---
threads_url: https://www.threads.com/@tin_ylab/post/DgDuhKOFRwN
blogger_url: https://tidytabs.blogspot.com/2026/11/how-to-split-shared-expenses-and-see.html
title: How to Split Shared Expenses and See Who Owes Whom in Excel and Google Sheets
labels: excel-formulas, small-business, expenses
tested_in: LibreOffice Calc 24.2.7.2 (Linux): template recalculated and compared with an independent Python model, 54 workbooks and 13,338 cell comparisons, 0 mismatches, plus the edge cases listed in the post. Not tested in Excel or Google Sheets.
search_description: Total what each partner paid with SUMIF, subtract a rounded fair share, and get settle lines like "Pay Alex $34.29" that add back to $0.00.
image_prompts: Three small piles of crumpled paper store receipts, one per partner, beside a zippered canvas cash pouch and a stack of folded tent weights on the open tailgate of a pickup truck at a weekend craft market, late afternoon light, realistic photo, no screens, no readable text, no logos
image_alt: Three piles of crumpled receipts beside a canvas cash pouch on a truck tailgate
threads: Who owes whom after a season of shared booth costs?\nTotal what each person paid with =SUMIF(paid_by,name,amount), subtract =ROUND(total*share,2), and the sign tells you.\nGive the leftover cent to one person, or three shares of $113.61 add up to $340.83 when you spent $340.82.
---
Total what each person paid with `=SUMIF(paid_by,name,amount)`, subtract their fair share `=ROUND(total*share,2)`, and the result is their balance: plus means they are owed, minus means they owe. In the made-up season of six booth expenses totaling $340.82, Alex paid $209.00, so with equal thirds the download writes "Pay Alex $34.29" for Bea and "Pay Alex $61.11" for Cal, and it puts the leftover cent on one row so the shares add back to $340.82 instead of $340.83.

> Works in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model (54 workbooks, 13,338 cell comparisons, 0 mismatches), plus the edge cases below. SUMIF, ROUND, SUMPRODUCT, COUNTIF and TEXT also exist in Excel and Google Sheets, but neither app was opened for this post.

![Six market expenses paid by Alex, Bea and Cal above balances and settle lines](images/2026-11-21-team-01-split-shared-expenses-who-owes-whom-sumif-round-excel-google-sheets/shared-expense-split-who-owes-whom-template.png) *Expenses tab with six fictional market costs totaling $340.82 above the People tab, where Alex paid $209.00 and collects $95.40, Bea pays Alex $34.29 and Cal pays Alex $61.11. PDF export from LibreOffice Calc 24.2.7.2.*

## Steps

The steps are the same in Excel and Google Sheets. Use the **Expenses** and **People** tabs.

1. On **Expenses**, type one shared expense per row starting in row 5: the date in **A** (09/05/2026), what it was in **B**, who paid in **C** and the amount in **D**. A refund is a negative amount; it is counted and lowers the total.
2. On **People**, type each name in **A5:A10**, up to six people. Row 5 is the treasurer, the one person everyone settles with.
3. Type each share in **B5:B10** as a percent. For equal thirds, type `=1/3` in each row so the shares add to exactly 100%. For uneven shares, type 50%, 30% and 20%.
4. In **B1**, total everything spent: `=SUM(Expenses!D5:D204)`.
5. In **C5**, total what that person paid: `=SUMIF(Expenses!$C$5:$C$204,A5,Expenses!$D$5:$D$204)`. Fill it down to row 10.
6. In **D6**, work out the fair share: `=ROUND($B$1*B6,2)`, the total times the share, rounded to the cent. Fill it down to **D10**.
7. In **D5**, give the treasurer the leftover cent: `=ROUND(ROUND($B$1*B5,2)+$B$1-SUMPRODUCT(ROUND($B$1*$B$5:$B$10,2)),2)`. SUMPRODUCT adds every rounded share, and the difference lands on this row.
8. In **E5**, get the balance: `=ROUND(C5-D5,2)`. Fill it down. The balances add to $0.00.
9. In **F6**, turn the balance into a sentence: `=IF(E6<0,"Pay "&$A$5&" "&TEXT(-E6,"$#,##0.00"),IF(E6>0,$A$5&" pays you "&TEXT(E6,"$#,##0.00"),"Even"))`. Fill it down to **F10**. On the treasurer row, **F5** shows the total they collect and pay out.

Settling through one person means at most (people - 1) payments. That keeps the sheet simple, but it is not always the fewest payments possible. The sheet does the arithmetic only: it does not decide how partners should split costs, and it does not cover how to record shared costs for a partnership or for taxes.

Microsoft's [SUMIF function](https://support.microsoft.com/en-us/office/sumif-function-169b8c99-c05c-4483-a712-1697a653039b) and [ROUND function](https://support.microsoft.com/en-us/office/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c) pages and Google's [SUMIF help page](https://support.google.com/docs/answer/3093583) cover the two core functions. The leftover-cent fix is explained in the Tidy Tabs posts on [splitting a shared booth fee by sales share](https://tidytabs.blogspot.com/2026/10/how-to-split-shared-booth-fee-between.html) and [fixing a total that is off by one cent](https://tidytabs.blogspot.com/2026/10/how-to-fix-total-that-is-off-by-one.html).

## Example

The partners and receipts are fictional. Three makers share a market booth for the fall season:

| Date | Description | Paid by | Amount |
|---|---|---|---|
| 09/05/2026 | Booth fee, fall season | Alex | $150.00 |
| 09/05/2026 | Tent weights | Bea | $45.99 |
| 09/12/2026 | Ice and water | Cal | $12.50 |
| 09/19/2026 | Card reader | Alex | $59.00 |
| 10/03/2026 | Signs and price tags | Bea | $33.33 |
| 10/17/2026 | Gas for the van | Cal | $40.00 |

Total spent: $340.82. With equal thirds, the **People** tab shows:

| Name | Share | Paid | Fair share | Balance | Settle line |
|---|---|---|---|---|---|
| Alex | 33.33% | $209.00 | $113.60 | $95.40 | Collects $95.40, pays out $0.00 |
| Bea | 33.33% | $79.32 | $113.61 | -$34.29 | Pay Alex $34.29 |
| Cal | 33.33% | $52.50 | $113.61 | -$61.11 | Pay Alex $61.11 |

Alex, the treasurer, takes the one-cent difference, so the fair shares add to $340.82 and the balances add to $0.00. Change the shares to 50%, 30% and 20% and the fair shares become $170.41, $102.25 and $68.16, with "Pay Alex $22.93" for Bea and "Pay Alex $15.66" for Cal.

## Troubleshooting

### The balances add to -$10.00 instead of $0.00
A receipt's payer does not match a name on **People**. In LibreOffice, a $10.00 napkin receipt paid by "Bea " (with a space at the end) was skipped by SUMIF, so Bea's paid total stayed $79.32 and **J9** showed $10.00 spent under names not on the list. The row's Status reads Check: payer not on People; retype the name or pick it from the dropdown. Upper and lower case did not matter in LibreOffice ("bea" counted for Bea).

### The fair shares add to $340.83
Each share was rounded on its own: three times $113.61 is $340.83, one cent more than was spent. Use the treasurer formula in **D5** so the leftover cent lands on one row.

### The Fair share column is blank
The shares do not add to 100%, or a People row shows Check. Shares of 99% or 101% leave **D:F** blank until they are fixed, and **J5** shows what they add to.

### A row shows Check: share (50% not 50)
The share was typed as 50 instead of 50%, so the shares add to 5050.00%. Retype it as a percent.

### An amount is not counted
The amount was typed as text, such as "$10.00". Status reads Check: amount is text and the amount stays out of the total. Retype it as a number.

## Template

[Download the .xlsx](templates/tidy-tabs-shared-expense-split.xlsx)

The file has three tabs. **Expenses** holds rows 5 to 204 with Date, Description, Paid by (a dropdown of the People names), Amount and a Status that flags text dates, missing or unknown payers, missing or text amounts, and counted refunds. Its summary in **G4:H9** counts rows, Checks, unknown payers, refunds and entries typed below row 204.

**People** holds up to six people in rows 5 to 10 with Name, Share %, Paid, Fair share, Balance, Settle line and Status, which catches extra spaces in names, names typed twice, shares typed as 50 instead of 50% and a blank treasurer row. The summary in **I4:J13** ends with a CHECK that reads OK only when the shares, totals and balances all tie out. **How to use** explains each formula. Both working tabs print Landscape on US Letter with the header row repeated.

Yellow cells are inputs and gray cells are formulas; every Check turns red. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
