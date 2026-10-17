---
threads_url: https://www.threads.com/@tin_ylab/post/DemE4IzkV7-
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-calculate-how-long-client-has.html
title: How to Calculate How Long a Client Has Been With You in Years and Months in Excel and Google Sheets
labels: excel-formulas, small-business, clients
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and every result was checked against Python month math, including a Feb 29 start, a start day later in the month than the as of day, a start date after the as of date, a date typed as text and a changed years threshold. Excel and Google Sheets were not opened; the help link is from Microsoft Support.
search_description: Client tenure in years and months: =DATEDIF(B2,$I$2,"y") and "ym". A 03/15/2021 client shows 5 yr 7 mo on 10/16/2026. Free .xlsx template.
image_prompts: Three worn manila project folders fanned out beside a brass desk bell and a small potted succulent on a freelancer's wooden desk, soft side light, no screens, no readable text, no logos
image_alt: Worn manila folders fanned out beside a brass bell and a succulent
threads: DATEDIF hides from the Excel function picker.\nType it in full anyway: =DATEDIF(B2,$I$2,"y") gives whole years and "ym" gives the months left over.\nA client who started 08/20/2025 shows 1 yr 1 mo on 10/16/2026, because the 20th has not arrived.
---
Count whole years with `=DATEDIF(B2,$I$2,"y")` and the leftover months with `=DATEDIF(B2,$I$2,"ym")`, where B2 is the first project date and I2 is the as-of date you type. Join them with `=C2&" yr "&D2&" mo"` to show text like 5 yr 7 mo.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Client list showing years and months with you for each first project date](images/2026-10-16-team-04-client-tenure-years-months-datedif-excel-google-sheets/client-tenure-years-months-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

DATEDIF counts complete years and months between two dates. A partial month does not count.

One catch for Excel: DATEDIF is not listed in the function picker (**Formulas > Insert Function**), and Excel does not suggest it as you type. It still works when you type it in full. Microsoft documents it on its [DATEDIF function](https://support.microsoft.com/en-us/office/datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c) page.

1. On a tab named **Clients**, type these headers in row 1: Client, First project, Years, Months, Tenure, Total months.
2. Type the client name in column **A** and the first project date as MM/DD/YYYY in column **B**.
3. In **I2**, type the as-of date, such as 10/16/2026. In **I3**, type the number of years that marks a long-term client, such as 3.
4. In **C2**, type `=IF(B2="","",IF(NOT(ISNUMBER(B2)),"Check date",IF(B2>$I$2,"Check date",DATEDIF(B2,$I$2,"y"))))`. This is whole years, or a warning if the date is not usable.
5. In **D2**, type `=IF(ISNUMBER(C2),DATEDIF(B2,$I$2,"ym"),"")`. This is the months left over after the whole years.
6. In **E2**, type `=IF(ISNUMBER(C2),C2&" yr "&D2&" mo","")`. In **F2**, type `=IF(ISNUMBER(C2),C2*12+D2,"")` for the total months, which sorts cleanly.
7. Fill C2:F2 down to row 30.
8. Add a summary. In **I4**, type `=COUNT(C2:C30)`. In **I5**, type `=COUNTIF(C2:C30,">="&I3)` for clients at or above your threshold. In **I7**, type `=COUNTIF(C2:C30,"Check date")`.

The IF wrappers matter. They catch a start date after the as-of date before DATEDIF runs, so you see "Check date" instead of an error. Microsoft's page says DATEDIF returns #NUM! when the start date is greater than the end date.

### In Excel

Type the formulas in full. Excel will not offer DATEDIF in a drop-down. Format column **B** and **I2** as dates from **Home > Number Format**.

### In Google Sheets

The same formulas should work as typed. Format dates with **Format > Number > Date**. This was not opened in Google Sheets for this post.

For a live count, put `=TODAY()` in I2. The sample uses the typed date 10/16/2026.

## Example

Sample data below is fictional. It is a freelancer's client list, counted as of 10/16/2026.

| Client | First project | Years | Months | Total months |
|---|---|---|---|---|
| Maple Street Bakery | 03/15/2021 | 5 | 7 | 67 |
| Harbor Yoga | 11/02/2022 | 3 | 11 | 47 |
| Birch Design | 02/29/2024 | 2 | 7 | 31 |
| Lantern Print | 08/20/2025 | 1 | 1 | 13 |
| Cedar Row | 10/16/2023 | 3 | 0 | 36 |
| Willow Pet | 06/30/2019 | 7 | 3 | 87 |
| Granite Fitness | 01/05/2026 | 0 | 9 | 9 |
| Juniper Books | 12/31/2020 | 5 | 9 | 69 |
| Copper Kettle | 10/01/2026 | 0 | 0 | 0 |
| Pine Hill Dental | 04/09/2018 | 8 | 6 | 102 |

| Result | Cell | Value |
|---|---|---|
| Clients counted | I4 | 10 |
| Clients at 3 years or more | I5 | 6 |
| Rows showing Check date | I7 | 0 |

Lantern Print started on the 20th. On 10/16/2026 the 20th has not arrived, so it shows 1 yr 1 mo, not 1 yr 2 mo.

## Troubleshooting

### A row says Check date
The start date is after the as-of date, or the date was typed as text. In the recalculated copy, a start of 01/01/2027 and a text entry of '03/15/2020' both showed Check date, and the check cell I8 still read OK. Fix the date or move I2 forward.

### I get #NUM! when I type DATEDIF without the check
Microsoft's [DATEDIF page](https://support.microsoft.com/en-us/office/datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c) says the result is #NUM! when the start date is greater than the end date. We did not run the bare formula on its own. The template avoids the case by testing the dates first.

### The months look one short
DATEDIF counts complete months. With a start of 08/20/2025 and an as-of date of 10/16/2026, the answer is 1 yr 1 mo because the 20th has not come yet. The total is 13 months.

## Template

[Download the .xlsx](templates/tidy-tabs-client-tenure.xlsx)

The file has two tabs. **Clients** holds the 10 sample clients in yellow input cells, room down to row 30, the as-of date in **I2**, the years threshold in **I3**, and the summary in **I4** to **I8**. **How to use** has short fill-in notes. Both tabs print portrait on US Letter.

The clients are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open the file in Google Sheets, go to **File > Import > Upload** and select it.
