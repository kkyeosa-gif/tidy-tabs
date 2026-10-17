---
title: How to Get the Domain or Username From an Email Address in Excel and Google Sheets
labels: excel-formulas, small-business, customers
tested_in: LibreOffice Calc 24.2 (Linux) only. The template was recalculated there and 28 fictional addresses were checked against Python string splitting, including two @ signs, no @, spaces, a missing dot and upper case. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Get the domain from an email with =MID(A2,FIND("@",A2)+1,LEN(A2)) in Excel and Google Sheets. Free email split template with a domain count.
image_prompts: Clipboard holding a lined paper sign-up sheet with illegible handwriting, a ballpoint pen on a string, and a small stack of kraft-paper wrapped soap bars on a farmers market table, soft daylight, no screens, no readable text, no logos
image_alt: Clipboard sign-up sheet and pen beside wrapped soap bars on a market table
threads: Need just the domain from a column of email addresses?\n=MID(A2,FIND("@",A2)+1,LEN(A2)) takes everything after the @, and LEFT with FIND takes the username.\nThe template also flags a trailing space, because "example.com " with a space is not the domain you think it is.
---
Find the @ sign, then take everything after it with `=MID(A2,FIND("@",A2)+1,LEN(A2))`. The username is everything before it: `=LEFT(A2,FIND("@",A2)-1)`.

> Works in: LibreOffice Calc 24.2 (Linux), where the template was built and every result below was recalculated and checked against Python. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Email list split into username and domain columns with a shape check](images/2026-10-18-team-03-extract-domain-from-email-address-left-mid-find-excel-google-sheets/email-domain-split-template.png) *caption: render of the template, PDF export from LibreOffice Calc 24.2, fictional sample data*

## Steps

Put the addresses in column **A**, starting in **A2**, with headers in row 1: Email as typed, Username, Domain.

1. Select column **A** and format it as Plain text before you paste, so the cells stay exactly as typed.
2. Paste or type the addresses into **A2** and down.
3. In **B2**, type `=LEFT(A2,FIND("@",A2)-1)`. `FIND` returns the position of the @, and `LEFT` keeps the characters before it. The `-1` leaves the @ out.
4. In **C2**, type `=MID(A2,FIND("@",A2)+1,LEN(A2))`. `MID` starts one character after the @ and `LEN(A2)` is a length long enough to reach the end of the address.
5. Select **B2:C2** and drag the fill handle down the list.

Those two formulas are the whole trick. A cell with no @ makes `FIND` return an error, so the template wraps them to keep the sheet clean:

- **B2**: `=IF(A2="","",IFERROR(LEFT(A2,FIND("@",A2)-1),""))`
- **C2**: `=IF(A2="","",IFERROR(MID(A2,FIND("@",A2)+1,LEN(A2)),""))`
- **D2**: `=IF(C2="","",LOWER(C2))` gives the domain in lower case, so Example.com and example.com count together.

6. To flag addresses that look wrong, type this in **E2** and fill it down:

`=IF(A2="","",IF(AND(LEN(A2)-LEN(SUBSTITUTE(A2,"@",""))=1,LEFT(A2,1)<>"@",ISNUMBER(FIND(".",A2,FIND("@",A2)+2)),RIGHT(A2,1)<>".",ISERROR(FIND(" ",A2))),"OK","Check"))`

It shows OK when the address has exactly one @, something before it, a dot at least one character after it, no dot at the very end and no spaces.

### In Excel

The formulas work as typed. Microsoft describes the position function on its [FIND function](https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628) page and the text slice on its [MID function](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028) page.

### In Google Sheets

The same formulas should work as typed. Google documents the position function on its [FIND function](https://support.google.com/docs/answer/3094126) page. None of this was opened in Google Sheets for this post.

## Example

Sample data below is fictional and uses the reserved example domains. It is a small soap shop cleaning up a sign-up list.

| Email as typed | Username | Domain | Domain (lower case) | Shape |
|---|---|---|---|---|
| Jane.Doe@Example.com | Jane.Doe | Example.com | example.com | OK |
| sam@shop.example.org | sam | shop.example.org | shop.example.org | OK |
| rosa.m+orders@example.net | rosa.m+orders | example.net | example.net | OK |
| BILL.R@EXAMPLE.ORG | BILL.R | EXAMPLE.ORG | example.org | OK |
| maria@@example.com | maria | @example.com | @example.com | Check |
| tomexample.com | | | | Check |
| ann lee@example.net | ann lee | example.net | example.net | Check |
| kim@example | kim | example | example | Check |

The template has 14 addresses in all. The **Summary** tab counts the OK rows per domain: example.com 3, example.org 2, example.net 2 and 2 other, with 9 OK and 5 to check. Fourteen more edge addresses matched Python as well.

## Troubleshooting

### The domain has an invisible space at the end
`pat@example.com ` with a trailing space gives the domain `example.com ` and a Check in **E**. A leading space also shows Check. Retype the cell, or run `TRIM` over the column and paste the results back as values. Spaces were tested; tabs and non-breaking spaces were not.

### Username and domain are both blank
There is no @, as in `tomexample.com`. Without the `IFERROR` wrapper, `FIND` returns an error here. With it, both cells stay blank and **E** says Check.

### A subdomain is counted under Other domains
`sam@shop.example.org` has the domain `shop.example.org`, which is not the same text as `example.org`. It lands in **Other domains** on the Summary tab. If you want it counted as example.org, type `shop.example.org` as its own line, or pull the last two parts of the domain yourself.

### A bad address shows OK
The shape check does not catch two dots in a row, so `two..dots@example.com` shows OK. It also does not check which characters are allowed or whether the ending is a real one. Treat OK as "looks like an address".

### Other domains is negative
You typed the same domain twice on the Summary tab, in different capitalization or with extra spaces, and it was counted twice. In the recalculated copy, `  EXAMPLE.COM ` in **A5** counted 3 like **A2**, and Other domains showed -1. Type each domain once.

## Limits

This checks the shape of the text only. It cannot tell you whether a mailbox exists, and it says nothing about whether the person agreed to get email from you. Rules for sending marketing email are a separate question, so check them before you use a list. The formulas run from row 2 to row 200.

## Template

[Download the .xlsx](templates/tidy-tabs-email-domain-split.xlsx)

The file has three tabs. **Emails** has the addresses in yellow input cells **A2:A60**, formatted as text, and the formulas for Username, Domain, Domain (lower case) and Shape in **B:E**, filled to row 200. Shape cells that say Check are colored. **Summary** holds the domains to count in yellow cells **A2:A7**, the OK counts beside them in **B2:B7**, and the totals and a check in **B8:B13**. **How to use** has short notes. All tabs print on US Letter, and the Emails print area is the first 40 rows.

The addresses are fictional. The file was recalculated only in LibreOffice Calc 24.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
