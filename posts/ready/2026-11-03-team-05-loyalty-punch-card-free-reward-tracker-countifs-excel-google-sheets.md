---
title: How to Track a Loyalty Punch Card and Free Rewards in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with an independent Python model over 14 workbooks and 1,448 comparisons with 0 mismatches. Checked: 0, 2 and 5 customers, 0, 10, 100 and 300 purchases, punches per reward of 1, 3, 5 and 10, a smallest purchase from $0.00 to $10.00, a customer typed in lower case, a customer with a trailing space, rewards used above rewards earned, and settings of 0, 2.5, text and a negative number. Excel and Google Sheets were not used.
search_description: Track a loyalty punch card in Excel and Google Sheets with COUNTIFS, INT and MOD. Free template shows punches, rewards ready and punches left.
image_prompts: A small stack of blank cardstock punch cards held with a binder clip beside a hand held hole punch and a ceramic coffee mug on a cafe counter, soft daylight, no screens, no readable text, no logos
image_alt: Stack of blank punch cards with a hole punch and a ceramic mug on a counter
threads: Punch card says 10 purchases = 1 free. Who has earned one?\nCOUNTIFS counts each customer's purchases of $5.00 or more. INT(punches/10) gives rewards earned, 10-MOD(punches,10) gives punches left.\nThe sheet also counts purchases typed under a name that is not on your list.
---
Count each customer's punches with `=COUNTIFS(Purchases!$B$5:$B$304,A6,Purchases!$C$5:$C$304,">="&$B$3)`, then get the free rewards earned with `=INT(B6/$B$2)`. The punches left to the next reward are `=$B$2-MOD(B6,$B$2)`.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. COUNTIFS, INT and MOD exist in current Excel and Google Sheets, but neither app was opened for this post.

![Customer table with punches, rewards earned, rewards used, rewards ready and punches to next reward](images/2026-11-03-team-05-loyalty-punch-card-free-reward-tracker-countifs-excel-google-sheets/loyalty-punch-card-tracker-template.png) *PDF export from LibreOffice Calc 24.2.7.2, fictional sample data*

## Steps

The steps are the same in Excel and Google Sheets. The layout matches the template. **Customers** has the settings in **B2** (punches for a reward, 10) and **B3** (smallest purchase that earns a punch, $5.00) and the customer table from row 6. **Purchases** has date, customer and amount from row 5.

1. On **Purchases**, log each sale: the date in A, the customer in B and the amount in C. Spell the customer exactly as it is on the Customers tab.
2. On **Customers**, list each customer once in **A6:A25**.
3. In **B6**, count the punches: `=COUNTIFS(Purchases!$B$5:$B$304,A6,Purchases!$C$5:$C$304,">="&$B$3)`. Only purchases of at least $5.00 count.
4. In **C6**, type `=INT(B6/$B$2)` for the rewards earned. INT drops the leftover punches.
5. In **D6**, type the rewards already given. In **E6**, type `=C6-D6` for the rewards ready.
6. In **F6**, type `=$B$2-MOD(B6,$B$2)`. MOD gives the leftover punches, so this is how many are left to the next reward.
7. Fill **B6:F6** down.

A customer with exactly 10 punches has 1 reward earned and 10 punches to the next one. The count starts over after each reward.

The sheet counts punches and rewards. It does not decide what a reward is or how long a punch lasts.

Help pages: Microsoft's [COUNTIFS function](https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842) and [INT function](https://support.microsoft.com/en-us/office/int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef), and Google's [function list](https://support.google.com/docs/table/25273). They describe the functions, not our test. Each link answered with HTTP 200 on 11/02/2026; the pages were not read for this post.

## Example

Sample data is fictional: eight customers of a small cafe, 87 logged purchases, 10 punches for a reward.

| Customer | Punches | Rewards earned | Rewards used | Rewards ready | Punches to next |
|---|---|---|---|---|---|
| Maya R. | 10 | 1 | 0 | 1 | 10 |
| Devon T. | 3 | 0 | 0 | 0 | 7 |
| Priya S. | 17 | 1 | 1 | 0 | 3 |
| Sam K. | 8 | 0 | 0 | 0 | 2 |
| Luis P. | 14 | 1 | 1 | 0 | 6 |

Maya R. shows Reward ready. Priya S. has 23 purchases logged, but only 17 were $5.00 or more, so she has 17 punches.

## Troubleshooting

### A customer has fewer punches than expected
The purchases under their name may be below the smallest amount in **B3**, or a name is spelled a different way. **J7** counts purchases whose customer is not on the Customers tab.

### J7 is above 0
A purchase has a name that is not in **A6:A25**, such as `Bo ` with a trailing space or a misspelling. In the test, upper and lower case did not matter (`ann` matched `Ann`), but a trailing space made a different customer. Fix the spelling.

### A status says Check
Rewards used is larger than rewards earned. Either the used number was mistyped or purchases are missing. In the test, the row showed Check and the **CHECK** cell said Check.

### Every row is blank and B4 shows a message
A setting is out of range. Punches for a reward must be a whole number of 1 or more and the smallest purchase 0 or more. In the test, 0, 2.5, text and a negative number each blanked the table with a message and no #DIV/0!.

### A customer is not on the list
Add them in the next empty cell of **A6:A25**. The table holds 20 customers and the log holds 300 purchases. Rows beyond that were not counted and were not tested.

## Template

[Download the .xlsx](templates/tidy-tabs-loyalty-punch-card.xlsx)

The file has three tabs. **Customers** has the two settings, a settings check in **B4**, the table for 20 customers, a summary and a **CHECK** cell, **J9**. **Purchases** has room for 300 sales with 87 fictional rows. **How to use** has notes on each formula.

**CHECK** says OK when the settings are valid, no purchase is unmatched and no row shows Check. Fix every Check before you give a reward.

To open it in Google Sheets, go to **File > Import > Upload** and select the file.
