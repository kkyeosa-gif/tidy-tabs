---
threads_url: https://www.threads.com/@tin_ylab/post/DeoprKJGxF_
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-format-phone-numbers-in-excel.html
title: How to Format Phone Numbers in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and 30 phone entries (12 sample rows plus 18 edge cases, such as extra spaces, a leading 1, an extension, letters and a slash) were compared with a Python clean-up. Excel and Google Sheets were not opened; the help links are from Microsoft Support.
search_description: Format phone numbers as (555) 123-4567 with ="("&LEFT(C2,3)&") "&MID(C2,4,3)&"-"&RIGHT(C2,4) in Excel and Google Sheets.
image_prompts: Wooden card file box with several handwritten index cards pulled halfway out beside a corded desk phone and a small sticky note pad, soft window light, no screens, no readable text or digits, no logos
image_alt: Index card box with handwritten cards next to a corded desk phone and sticky notes
threads: Phone numbers typed five different ways make a mess of a contact list.\nNested SUBSTITUTE strips the punctuation, then LEFT, MID and RIGHT rebuild (555) 123-4567.\nAnything that is not 10 digits shows Check instead of a wrong guess.
---
Strip spaces, hyphens, dots, parentheses and plus signs with nested SUBSTITUTE, then rebuild the 10 digits in C2 (the version without a leading 1) with `=IF(AND(LEN(C2)=10,SUMPRODUCT(--ISNUMBER(FIND(MID(C2,{1,2,3,4,5,6,7,8,9,10},1),"0123456789")))=10),"("&LEFT(C2,3)&") "&MID(C2,4,3)&"-"&RIGHT(C2,4),"Check")`. A number like 5551234567 becomes (555) 123-4567, and anything that is not 10 digits shows Check.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Phone list with the digits-only and formatted columns and Check flags](images/2026-10-17-team-04-format-phone-numbers-substitute-left-mid-right-excel-google-sheets/format-phone-numbers-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

The sheet has the phone as typed in **A**, digits only in **B**, the number without a leading 1 in **C** and the formatted number in **D**.

1. On a tab named **Phones**, type these headers in row 1: Phone as typed, Digits only, Without leading 1, Formatted.
2. Format columns **A** to **C** as plain text before you paste numbers, so a leading zero is not dropped.
3. In **B2**, type `=IF(A2="","",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2," ",""),"-",""),"(",""),")",""),".",""),"+",""))`. This removes the six characters people use as separators.
4. In **C2**, type `=IF(B2="","",IF(AND(LEN(B2)=11,LEFT(B2,1)="1"),MID(B2,2,10),B2))`. This drops a leading 1 from an 11 digit number.
5. In **D2**, type `=IF(A2="","",IF(AND(LEN(C2)=10,SUMPRODUCT(--ISNUMBER(FIND(MID(C2,{1,2,3,4,5,6,7,8,9,10},1),"0123456789")))=10),"("&LEFT(C2,3)&") "&MID(C2,4,3)&"-"&RIGHT(C2,4),"Check"))`.
6. Select **B2:D2** and drag the fill handle down to row 40.
7. For a summary, type the labels Rows with a phone entry, Formatted, Rows to check and Check: formatted + check = rows in **F2:F5**. Then put `=COUNTA(A2:A40)` in **G2**, `=SUMPRODUCT(--(LEN(D2:D40)=14))` in **G3** and `=COUNTIF(D2:D40,"Check")` in **G4**.
8. In **G5**, type `=IF(G3+G4=G2,"OK","CHECK")`. Formatted rows plus Check rows should equal the rows with data.

Then filter column **D** to Check and fix those rows by hand.

The long test in step 5 confirms that all 10 characters are digits. The shorter `ISNUMBER(--C2)` is not enough, because `--"5551e34567"` reads as a number in scientific notation and would be formatted as a phone number.

### In Excel

The formulas work as typed. Microsoft describes the pieces on its [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332) and [LEFT, LEFTB functions](https://support.microsoft.com/en-us/office/left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c) pages. To filter, the usual menu path is **Data > Filter**; Microsoft explains it on its [Filter data in a range or table](https://support.microsoft.com/en-us/office/filter-data-in-a-range-or-table-01832226-31b5-4568-8806-38c37dcc180e) page.

### In Google Sheets

The same formulas should work as typed, and the **Data > Create a filter** menu item should give you the same Check view. Neither was opened in Google Sheets for this post.

## Example

Sample data below is fictional. All numbers use the reserved 555-01xx style.

| Phone as typed | Formatted |
|---|---|
| (212) 555-0101 | (212) 555-0101 |
| 212.555.0102 | (212) 555-0102 |
| +1 212-555-0103 | (212) 555-0103 |
| 2125550104 (typed as a number) | (212) 555-0104 |
| 1-212-555-0106 | (212) 555-0106 |
| (415)555-0107 | (415) 555-0107 |
| 212-555-010 | Check |
| 212-555-0111 x12 | Check |
| call after 5 | Check |

The table shows 9 of the 12 sample rows. The full list gives G2 = 12, G3 = 8, G4 = 4 and G5 = OK.

## Troubleshooting

### An extension shows Check
A number such as 212-555-0111 x12 has the letter x and two extra digits, so it is not 10 digits. The template does not split extensions. One idea, not run in the template: move the extension to its own column by hand, then format the rest.

### A number with a slash shows Check
In the checked copy, 555/123/4567 showed Check because the chain does not strip slashes. If your list uses slashes, one idea is to wrap one more `SUBSTITUTE(...,"/","")` around the others in **B2**. That change was not run.

### A number from another country shows Check
+44 20 7946 0958 loses its plus and spaces, but the result is 12 digits, so D shows Check. The same happens to 11 digits that do not start with 1. The template handles US 10 digit numbers, with an optional leading 1.

### A leading zero disappeared
A number typed as plain 0551234567 into a number cell becomes 551234567, which is 9 digits and shows Check. As text, 0551234567 was formatted as (055) 123-4567. Format column **A** as plain text before you paste.

### Two numbers look identical but one shows Check
A non-breaking space, character 160, is not removed by `SUBSTITUTE(A2," ","")`, which only matches a regular space. In the checked copy, a number with one inside showed Check. The [non-breaking space template](templates/tidy-tabs-remove-non-breaking-spaces.xlsx) on this blog cleans it first with `=TRIM(CLEAN(SUBSTITUTE(A2,UNICHAR(160)," ")))` in its Cleaned column; that template turned `Dana` + character 160 + `Ruiz` into Dana Ruiz in LibreOffice Calc.

## Limits

This only changes the format. It does not check that a number exists or can be called. Extensions, letters and other countries show Check, and slashes and non-breaking spaces are not stripped. The formulas stop at row 40, so extend them for a longer list.

## Template

[Download the .xlsx](templates/tidy-tabs-format-phone-numbers.xlsx)

The file has two tabs. **Phones** holds 12 sample rows in yellow input cells in **A2:A40**, formulas in **B** to **D** filled to row 40, the summary labels in **F2:F5** with the results in **G2:G5**. **How to use** has short fill-in notes. Columns **A** to **C** are formatted as text.

The phone numbers are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
