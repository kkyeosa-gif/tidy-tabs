---
threads_url: https://www.threads.com/@tin_ylab/post/DeO5uVQoNm_
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-freelance-hourly-rate.html
title: How to Calculate a Freelance Hourly Rate in Excel and Google Sheets
labels: excel-formulas, freelance, small-business-finance
tested_in: LibreOffice Calc 24.2.7 (Linux). Excel and Google Sheets were not opened; the ROUNDUP help pages are linked.
search_description: Calculate a freelance hourly rate with =B10/B9: income goal plus overhead divided by billable hours. Free .xlsx template with three what-if columns.
image_prompts: A freelancer home studio corner with a paper wall calendar with a few days circled, a stack of sketchbooks, and a basic desk calculator beside a pencil cup, soft daylight, all calendar and paper text blurred and unreadable, no screens, no logos
image_alt: Wall calendar with circled days above sketchbooks, a pencil cup, and a desk calculator
threads: Dividing your income goal by every hour you work hides the unpaid ones.\nDivide by billable hours instead: $66,000.00 over 936 hours is $70.51 an hour. Over all 1,440 hours it would look like $45.83.
---
Divide what you need to bring in by the hours you can bill: `=B10/B9`. A $60,000.00 goal plus $6,000.00 overhead over 936 billable hours comes to $70.51 an hour.

> Works in: LibreOffice Calc 24.2.7 (Linux), where the template was built and every number below was recalculated. The formulas use only basic arithmetic and ROUNDUP, which Excel and Google Sheets both have, but neither app was opened for this post.

![Hourly rate sheet comparing the base case with three what-if columns and the resulting rates](images/2026-10-07-team-04-freelance-hourly-rate-calculator-excel-google-sheets/freelance-hourly-rate-calculator-template.png) *Template with sample data (fake), PDF export from LibreOffice Calc 24.2.*

This is an illustration of the math, not tax or financial advice. It leaves out taxes, self-employment costs and benefits, so ask an accountant about those before you set a price.

## Steps

Type labels in column A and values in column B. The row numbers below match the template.

1. In **A1**, type `Your numbers` as a heading for column B.
2. In **B2**, type your yearly income goal: 60000.
3. In **B3**, type your yearly overhead (software, supplies, a co-working pass): 6000.
4. In **B4**, type the weeks off per year, including holidays and sick days: 4.
5. In **B5**, type the hours per week you work: 30.
6. In **B6**, type the share of those hours you can bill, as a percent: 65%.
7. In **B7**, type `=52-B4` for weeks worked.
8. In **B8**, type `=B7*B5` for hours worked per year.
9. In **B9**, type `=B8*B6` for billable hours. This is the number that matters most.
10. In **B10**, type `=B2+B3` for the money you need to bring in.
11. In **B11**, type `=B10/B9` for your hourly rate.
12. In **B12**, type `=ROUNDUP(B11,0)` if you want a whole-dollar rate, such as $71.
13. Format B2, B3, B10, B11 and B12 as currency, and B6 as a percentage.

Billable hours are not all your hours. Finding clients, sending invoices and answering email take time you cannot charge for, so the sheet multiplies by a billable share (step 6) before it divides.

### In Excel

The formulas work as typed. Microsoft documents the rounding step on its [ROUNDUP function](https://support.microsoft.com/en-us/office/roundup-function-f8bc9b23-e795-47db-8703-db171d0c42a7) page.

### In Google Sheets

The formulas work as typed. Google documents the rounding step on its [ROUNDUP](https://support.google.com/docs/answer/3093443) help page.

## Example

Sample data below is fictional. It is one freelancer, then three what-if columns that change one input each.

| | Your numbers | Higher goal | More time off | More billable |
|---|---|---|---|---|
| Yearly income goal | $60,000.00 | $75,000.00 | $60,000.00 | $60,000.00 |
| Yearly overhead | $6,000.00 | $6,000.00 | $6,000.00 | $6,000.00 |
| Weeks off | 4 | 4 | 6 | 4 |
| Hours per week | 30 | 30 | 30 | 30 |
| Billable share | 65% | 65% | 65% | 75% |
| Weeks worked | 48 | 48 | 46 | 48 |
| Hours worked | 1,440 | 1,440 | 1,380 | 1,440 |
| Billable hours | 936 | 936 | 897 | 1,080 |
| Money needed | $66,000.00 | $81,000.00 | $66,000.00 | $66,000.00 |
| Hourly rate | $70.51 | $86.54 | $73.58 | $61.11 |

The first column checks out by hand. 52 - 4 = 48 weeks, 48 x 30 = 1,440 hours, and 1,440 x 65% = 936 billable hours. $66,000.00 / 936 = 70.5128, which shows as $70.51.

Compare the columns. Two more weeks off raises the rate by $3.07, because you have fewer hours to spread the same money over. Raising the billable share from 65% to 75% lowers the rate to $61.11, because more of your week earns money.

## Troubleshooting

### The rate is far lower than expected
The formula divides by total hours instead of billable hours. Point it at the billable hours cell (**B9**), not hours worked (**B8**). With 1,440 hours, the sample rate would drop to $45.83, which hides the unpaid time.

### The billable share shows as 0.65 or 6500%
The cell format and the typed value do not match. Format B6 as a percentage, then type 65% or 0.65. Typing 65 in a cell already formatted as a percent gives 6500%.

### The rate shows as a long decimal
The cell is formatted as General or Number. Format B11 as currency to show two decimals. The value underneath is still 70.5128, so rounding for display does not change later math.

### The rounded-up rate is a dollar higher than you expect
ROUNDUP always rounds away from zero, so $70.01 becomes $71. Use `=ROUND(B11,0)` instead if you want $70.51 to become $71 and $70.40 to become $70.

### A what-if column does not change
Its inputs are typed over the formulas. Each what-if column in the template has the same formulas as column B, so change only the five input rows (2 to 6) and leave rows 7 to 12 alone.

## Template

[Download the .xlsx](templates/tidy-tabs-freelance-hourly-rate-calculator.xlsx)

The file has two tabs. **Hourly rate** holds the five inputs, the six calculation rows from weeks worked to the rounded-up rate, and the three what-if columns. **How to use** has short notes on each row and repeats that the sheet does not cover taxes or self-employment costs.

All numbers in the file are made up. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
