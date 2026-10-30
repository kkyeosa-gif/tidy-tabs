---
threads_url: https://www.threads.com/@tin_ylab/post/DfGQA8Jo88B
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-label-customers-bronze-silver-or.html
title: How to Label Customers Bronze, Silver or Gold by Total Spend in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python threshold function. Checked: spend exactly on a cutoff, one cent under a cutoff, zero, a blank, a negative, text typed as spend, cutoffs in the wrong order, equal, zero or text, different cutoffs, 100 customers in two random sets, and a test formula without the TRUE catch-all. The IFS formula was checked only as the _xlfn.IFS form written into the .xlsx file; IFS typed directly in the app was not run. The nested IF version matched IFS in every case. Excel and Google Sheets were not used.
search_description: Label customers Bronze, Silver or Gold by total spend in Excel and Google Sheets with IFS. Free template with cutoffs you set, dollars to the next level and a count per label.
image_prompts: A small stack of three plain kraft paper gift tags in bronze, silver and gold ink tones beside a brass hole punch and a spool of twine on a wooden counter, soft daylight, no screens, no readable text, no logos
image_alt: Three paper gift tags in bronze, silver and gold tones beside a hole punch
threads: Bronze, Silver or Gold from one spend column, with cutoffs you can change.\nIFS reads top down: Gold first, then Silver, and a closing TRUE,"Bronze" so nobody gets #N/A.\nThere is also a column for dollars to the next level, and Gold stays blank there.
---
Label each customer with `=IFS(C5>=Cutoffs!$B$3,"Gold",C5>=Cutoffs!$B$2,"Silver",TRUE,"Bronze")`, where C5 is lifetime spend, B3 is the Gold cutoff and B2 is the Silver cutoff. The `TRUE` at the end is the catch-all that returns Bronze instead of #N/A.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every label and number below was recalculated and compared with a Python model. IFS was tested only in the form saved inside the .xlsx file (`_xlfn.IFS`). The nested IF version in this post gave the same result as IFS in every test. Neither Excel nor Google Sheets was opened for this post.

![Twenty-five customers with spend and Bronze, Silver or Gold labels, plus a count per label](images/2026-10-29-team-05-label-customers-bronze-silver-gold-by-spend-ifs-excel-google-sheets/customer-tiers-bronze-silver-gold-template.png) *caption: PDF export/render of the template with fictional sample data from LibreOffice Calc 24.2.7.2*

## Steps

Keep the two cutoffs in cells, not typed inside the formula. When you change a cutoff, every label updates.

The layout below matches the template. The **Cutoffs** sheet holds Silver in **B2** and Gold in **B3**. The **Customers** sheet has the customer name in column A, city in B and lifetime spend in C, starting in row 5.

IFS checks its pairs from left to right and stops at the first one that is true. So test the highest level first. A $1,240.50 customer passes the Silver test too, but the Gold test comes first and wins.

### In Excel

1. On the **Cutoffs** sheet, type 500 in **B2** (Silver from) and 1000 in **B3** (Gold from).
2. On **Customers**, type a name in **A5** and lifetime spend in **C5**, such as 1240.50. Add more customers below.
3. In **D5**, type `=IFS(C5>=Cutoffs!$B$3,"Gold",C5>=Cutoffs!$B$2,"Silver",TRUE,"Bronze")`.
4. Drag the fill handle in **D5** down to your last customer.
5. In **E5**, type `=IF(D5="Gold","",IF(D5="Silver",Cutoffs!$B$3-C5,Cutoffs!$B$2-C5))`. This is the dollars to the next label. Gold stays blank.
6. Drag the fill handle in **E5** down. Format the column as currency.

Microsoft's page: [IFS function](https://support.microsoft.com/en-us/office/ifs-function-36329a26-37b2-467c-972b-4a39bd951d45). That page describes the function, not our test. This post does not say which Excel versions include IFS. If your copy does not know IFS, use the nested IF below.

### In Google Sheets

1. Enter the same cells and formulas as in steps 1 to 6 above, typed exactly as written.
2. Fill **D5** and **E5** down the same way.

Google's help page: [IFS function](https://support.google.com/docs/answer/7014145). That page describes the function, not our test.

### Why the last pair is TRUE

IFS has no "otherwise" slot. If no pair is true, it returns #N/A. Ending with `TRUE,"Bronze"` makes the last pair always true, so anyone under Silver gets Bronze. In a test formula without that pair, a spend below Silver gave #N/A.

### The same label without IFS

Nested IF gives the same labels and works anywhere IF works:

`=IF(C5>=Cutoffs!$B$3,"Gold",IF(C5>=Cutoffs!$B$2,"Silver","Bronze"))`

It tests Gold first, then Silver, and everything else is Bronze. In the tests, this column matched the IFS column on every row. Microsoft's page: [IF function](https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2).

### The version in the template

The template's D5 wraps the IFS in guards so a bad entry shows a message instead of a wrong label:

`=IF(C5="","",IF(Cutoffs!$B$4<>"OK","Check cutoffs",IF(OR(NOT(ISNUMBER(C5)),C5<0),"Check spend",IFS(C5>=Cutoffs!$B$3,"Gold",C5>=Cutoffs!$B$2,"Silver",TRUE,"Bronze"))))`

The `ISNUMBER` guard matters. In the LibreOffice test, text such as "abc" passed the Gold test, so without the guard it was labeled Gold. In the file, the formula is stored as `_xlfn.IFS`, which is how LibreOffice read it. Typing IFS directly in the app is a different path that was not run.

The dollars-to-next-level cell is `=IF(OR(D5="Silver",D5="Bronze"),ROUND(IF(D5="Silver",Cutoffs!$B$3,Cutoffs!$B$2)-C5,2),"")`. The ROUND removes tiny float tails after the subtraction.

**Cutoffs B4** checks the two cutoffs. It says OK when Silver is above 0 and below Gold.

## Example

Sample data is fictional: 25 made up customers with lifetime spend in dollars, Silver from $500.00 and Gold from $1,000.00.

| Customer | Lifetime spend | Label | To next label |
|---|---|---|---|
| Alma R. | $1,240.50 | Gold | |
| Ben T. | $85.00 | Bronze | $415.00 |
| Cora L. | $500.00 | Silver | $500.00 |
| Dev K. | $499.99 | Bronze | $0.01 |
| Elena M. | $1,000.00 | Gold | |
| Hugo P. | $0.00 | Bronze | $500.00 |
| Jonah B. | $975.40 | Silver | $24.60 |
| Opal F. | $999.99 | Silver | $0.01 |
| Wren I. | $499.00 | Bronze | $1.00 |

The summary beside the list uses `=COUNTIF(D5:D104,H5)` and `=SUMIF(D5:D104,H5,C5:C104)`, with the label name in column H. For all 25 customers: Gold is 6 customers and $8,945.75, Silver is 8 customers and $5,811.74, and Bronze is 11 customers and $2,367.99. The check cell says OK.

Spend exactly on a cutoff gets the higher label: $500.00 is Silver and $1,000.00 is Gold. One cent under stays in the lower label. Dev K. at $499.99 is Bronze and needs $0.01 more.

This sorts customers by the cutoffs you type. It does not say what any label should earn a customer.

## Troubleshooting

### The label shows #N/A
The IFS formula has no `TRUE,"Bronze"` pair at the end, and the spend is below the Silver cutoff. Add the pair.

### The label shows #NAME?
Your app does not recognize IFS. Use the nested IF formula above instead. In LibreOffice, a file that stored IFS without the `_xlfn.` prefix also gave #NAME?. The template file stores it correctly.

### Every row says Check cutoffs
The two cutoffs are out of order, equal, zero, negative or text. **Cutoffs B4** shows the reason. Type numbers so that Silver is above 0 and below Gold. In the tests, no Gold, Silver or Bronze counts appeared until the cutoffs were fixed.

### One row says Check spend
The spend is text, such as "abc" or the text "750", or it is negative. Retype it as a positive number. The row is counted in the "rows that need a look" cell, and the check cell stays OK when that is the only problem.

## Template

[Download the .xlsx](templates/tidy-tabs-customer-tiers.xlsx)

The file has three tabs. **Customers** holds 100 customer slots, with yellow inputs in **A5:C104** and 25 fictional rows filled in. Gray cells are formulas: the IFS label in **D**, dollars to the next label in **E**, the nested IF label in **F**, and the summary in **H** to **J**. **Cutoffs** holds the two yellow cutoffs in **B2** and **B3** and the cutoff check in **B4**. **How to use** has short notes.

Gold, Silver and Bronze rows are colored, and Check messages are red. The file was recalculated only in LibreOffice Calc 24.2.7.2. It prints on US Letter, landscape for the Customers tab. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
