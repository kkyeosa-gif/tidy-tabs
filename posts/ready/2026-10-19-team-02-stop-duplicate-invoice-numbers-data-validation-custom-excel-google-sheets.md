---
title: How to Stop Duplicate Invoice Numbers and Bad Dates From Being Typed in Excel and Google Sheets
labels: excel-formulas, small-business, data-validation
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and the Rules check formulas (the same tests the validation rules use) were compared with a Python model on 12 sample rows and about 10 edited variants, with no error values. The three rules were read back from the saved .xlsx file. The live Stop alert and the input prompts were not exercised in any app, because headless LibreOffice does not evaluate them. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Stop duplicate invoice numbers with the custom data validation rule =COUNTIF($A$2:$A$200,A2)=1 in Excel and Google Sheets. Free validation rules template.
image_prompts: Stack of carbon-copy invoice books with two pages pulled out and laid side by side, a rubber date stamp and an ink pad on a small shop counter, soft natural window light, no screens, no readable text, no logos
image_alt: Duplicate invoice pages laid side by side next to a date stamp and ink pad
threads: Two invoices both numbered 1042 is a bad afternoon.\nA Custom validation rule, =COUNTIF($A$2:$A$200,A2)=1, can refuse the repeat. Template has a check tab that lists the rows breaking the rules.
---
Block repeated invoice numbers with a Custom data validation rule on the invoice column: `=COUNTIF($A$2:$A$200,A2)=1`. For dates, use `=AND(ISNUMBER(B2),B2>=DATE(2026,1,1),B2<=DATE(2026,12,31))` on the date column.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and the rule formulas were recalculated and checked against Python. The rule text was read back from the saved file. Typing a bad value and seeing the Stop alert was not tried in any app, and neither Excel nor Google Sheets was opened for this post.

![Twelve sample invoices on the Entries tab and a Rules check tab flagging three rows](images/2026-10-19-team-02-stop-duplicate-invoice-numbers-data-validation-custom-excel-google-sheets/validation-rules-invoice-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The list sits on a tab named **Entries**: Invoice # in **A**, Date in **B** and Qty in **C**, with headers in row 1 and data from row 2. Each rule is written for the first cell of its column. The references without `$` shift down one row for every cell in the range.

### In Excel

1. Select **A2:A200**.
2. Choose **Data > Data Validation**. On the **Settings** tab, set **Allow** to **Custom**.
3. In **Formula**, type `=COUNTIF($A$2:$A$200,A2)=1`. Leave **Ignore blank** checked.
4. On the **Error Alert** tab, set **Style** to **Stop** and type a title and message. Click **OK**.
5. Select **B2:B200** and repeat steps 2 to 4 with `=AND(ISNUMBER(B2),B2>=DATE(2026,1,1),B2<=DATE(2026,12,31))`.
6. Select **C2:C200** and repeat with `=IF(ISNUMBER(C2),AND(C2>0,C2=INT(C2)),FALSE)`. It accepts whole numbers above 0 only.

Microsoft documents the dialog on its [apply data validation to cells](https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) page, and `COUNTIF` on its [COUNTIF function](https://support.microsoft.com/en-us/office/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34) page.

### In Google Sheets

1. Select **A2:A200** and choose **Data > Data validation**.
2. Add a rule, set the criteria to **Custom formula is**, and type `=COUNTIF($A$2:$A$200,A2)=1`.
3. Set the rule to reject invalid input if the dialog offers that choice, then save.
4. Repeat for **B2:B200** and **C2:C200** with the formulas above.

Google's steps are on its [data validation help page](https://support.google.com/docs/answer/186103?hl=en). How Sheets treats the Stop style and invalid pasted data was not tested.

### Why the quantity rule uses IF

Do not write the quantity rule as `AND(ISNUMBER(C2),C2>0,C2=INT(C2))`. `AND` evaluates every argument, and `INT` of text is an error, so a typed word turns the whole test into an error instead of FALSE. The `IF(ISNUMBER(C2), ... ,FALSE)` form returns FALSE for text.

## Example

Sample data below is fictional. Twelve invoices were typed on purpose with two problems: INV-1006 appears twice, and INV-1008 has a quantity of 2.5.

| Entries row | Invoice # | Date | Qty | Rules check result |
|---|---|---|---|---|
| 2 | INV-1001 | 10/01/2026 | 4 | OK |
| 3 | INV-1002 | 10/02/2026 | 1 | OK |
| 4 | INV-1003 | 10/02/2026 | 12 | OK |
| 5 | INV-1004 | 10/05/2026 | 3 | OK |
| 6 | INV-1005 | 10/06/2026 | 6 | OK |
| 7 | INV-1006 | 10/07/2026 | 2 | Fix |
| 8 | INV-1007 | 10/08/2026 | 8 | OK |
| 9 | INV-1008 | 10/09/2026 | 2.5 | Fix |
| 10 | INV-1009 | 10/12/2026 | 5 | OK |
| 11 | INV-1006 | 10/13/2026 | 1 | Fix |
| 12 | INV-1010 | 10/14/2026 | 20 | OK |
| 13 | INV-1011 | 10/15/2026 | 7 | OK |

The **Rules check** tab shows 12 rows, 9 OK and 3 to fix. Two rows fail the invoice rule, none fail the date rule and one fails the quantity rule. A repeated number fails on both rows, which is why two bad entries give three rows to fix.

## Troubleshooting

### A duplicate or bad value was pasted in and nothing stopped it
Validation checks typing, not pasting. Pasting over a cell skips the rule, and values already in the cells are never checked. Look at the **Rules check** tab for Fix rows.

### inv-1001 is rejected but INV-1001 with a trailing space is accepted
`COUNTIF` ignores upper and lower case, so inv-1001 repeats INV-1001. It does not ignore spaces, so `INV-1001 ` with a space at the end counts as a different number. In the edited copy that row passed the invoice test.

### A real 2026 date is rejected
The date rule accepts whole days from 01/01/2026 to 12/31/2026. A date typed as text, such as one that stays left aligned, fails because of `ISNUMBER`. A date with a time on 12/31/2026 fails too, because 12/31/2026 3 PM is later than `DATE(2026,12,31)`. Change the two `DATE` values to move the year.

### Invoice numbers with * or ? behave oddly
`COUNTIF` treats a star or question mark in an invoice number as a wildcard. That was not tested, so avoid those characters in invoice numbers.

## Limits

The checks cover rows 2 to 200 and the date rule is fixed to 2026. This is a typing guard, not bookkeeping, and it gives no record-keeping advice.

## Template

[Download the .xlsx](templates/tidy-tabs-validation-rules.xlsx)

The file has three tabs. **Entries** holds the 12 sample rows in yellow cells **A2:C60**, with the three validation rules, input prompts and Stop alerts set on **A2:C200**. **Rules check** repeats each rule as TRUE or FALSE per row, adds a Row result of OK or Fix, and counts in **H2:H9**, where **H9** says OK when the counts agree. **How to use** has short notes. Each tab prints portrait on US Letter, first 40 rows only.

The invoices, dates and quantities are fictional. The file was recalculated only in LibreOffice Calc 24.2, and the Stop alert was not tried by typing. To open it in Google Sheets, go to **File > Import > Upload** and select the file.