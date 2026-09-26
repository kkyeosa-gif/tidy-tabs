---
threads_url: https://www.threads.com/@tin_ylab/post/DduRg0knwuj
blogger_url: https://tidytabs.blogspot.com/2026/09/freelance-project-tracker-in-google.html
title: Freelance Project Tracker in Google Sheets (Free Template)
labels: google-sheets-templates, freelance
tested_in: LibreOffice Calc 24.2.7 (Linux), 09/25/2026; Excel and Google Sheets steps from Microsoft/Google help pages.
image_prompts: A freelancer working at a home desk with a laptop showing a project tracker spreadsheet with green and red status cells, screen slightly blurred, coffee mug and notebook beside it
image_alt: Freelancer at a home desk with a project tracker spreadsheet open
search_description: Free freelance project tracker for Google Sheets: one row per project, with Amount, Days left, and Payment status calculated by formulas.
threads: Which client still owes you money? If the answer lives in your inbox, try one row per project instead.\nThis free Google Sheets tracker works out Amount, Days left, and Payment status on its own.\nAn invoice turns Overdue 30 days after you send it. Change the 30 if your terms are net 15.
---
This free tracker gives you one row per client project, and the Amount, Days left, and Payment status columns calculate themselves. Download the .xlsx below and bring it into Google Sheets with **File > Import**, or open it directly in Excel.

> Works in: Google Sheets (import steps from Google Docs Editors Help) and Excel for Microsoft 365 (open the .xlsx). Tested in: LibreOffice Calc 24.2.7 on Linux, 09/25/2026. We have not checked it in Excel or Google Sheets ourselves.

![Freelance project tracker with due dates, status, amounts, and payment status for five clients](images/2026-09-30-a-freelance-project-tracker-google-sheets/freelance-project-tracker-template.png) *The tracker in this post with sample projects (some columns hidden). PDF export from LibreOffice Calc 24.2.*

## Steps

### Import it into Google Sheets
1. Download the template from the Template section below.
2. Open a Google Sheets spreadsheet and click **File > Import**.
3. Choose the .xlsx file from your computer.
4. Under import location, pick **Replace spreadsheet** (use a blank spreadsheet) or **Insert new sheets** (add the tabs to a file you already have).
5. Click **Import**.

Google lists every import option on its [import help page](https://support.google.com/docs/answer/40608).

### Fill in a project
1. On the **Projects** tab, type the Client, Project, Start date, and Due date in a new row. Type dates as MM/DD/YYYY.
2. Pick a Status from the dropdown: Not started, In progress, Waiting on client, or Done.
3. Pick Billing: **Hourly** or **Flat**.
4. For Hourly, fill Hours and Rate. For Flat, fill Flat fee. Amount picks the right one.
5. When you send the invoice, type that date in **Invoice sent**.
6. When the money arrives, type that date in **Paid on**.

Don't type over Amount, Days left, or Payment. Those are formulas.

Days left counts down to the due date. A negative number means the project is late, and the cell turns red. The number disappears once Status is Done.

Payment shows one of four words: Not invoiced, Waiting, Overdue, or Paid. Overdue means more than 30 days have passed since Invoice sent. Net 30 is a common invoice term, but it is our default, not a legal rule. If your terms are net 15, change the 30 in the formula to 15.

### How the Summary tab works
The **Summary** tab reads rows 2 to 200 of Projects and shows four numbers:

- **Open projects:** rows with a Client whose Status is not Done (`COUNTIFS`).
- **Billed, not paid ($):** total Amount where Invoice sent has a date and Paid on is blank (`SUMIFS`).
- **Paid ($):** total Amount where Paid on has a date.
- **Invoices overdue (30+ days):** how many rows show Overdue.

### Build it yourself
If you'd rather add these columns to your own sheet, here are the row 2 formulas from the template. Column F is Billing, G Hours, H Rate, I Flat fee, D Due date, E Status, K Invoice sent, L Paid on.

- Amount (J2): `=IF(F2="Hourly",G2*H2,I2)`
- Days left (M2): `=IF(OR(D2="",E2="Done"),"",D2-TODAY())`
- Payment (N2): `=IF(L2<>"","Paid",IF(K2="","Not invoiced",IF(TODAY()-K2>30,"Overdue","Waiting")))`

Copy each formula down the column.

To add the Status dropdown in Google Sheets:
1. Select E2:E200.
2. Click **Data > Data validation**, then **Add rule**.
3. Under criteria, choose **Dropdown** and type each option, clicking **Add another item** for the next one.
4. Click **Done**.

See Google's [Create an in-cell dropdown list](https://support.google.com/docs/answer/186103) for the other ways to start a dropdown.

To shade overdue invoices red:
1. Select N2:N200.
2. Click **Format > Conditional formatting**.
3. Under **Format cells if**, choose **Custom formula is** and type `=$N2="Overdue"`.
4. Pick a fill color and click **Done**.

Google's [conditional formatting help page](https://support.google.com/docs/answer/78413) covers custom formulas in more detail.

## Example

Here are the sample rows, as they calculated on 09/25/2026:

| Client | Project | Billing | Amount | Invoice sent | Paid on | Payment |
| --- | --- | --- | --- | --- | --- | --- |
| Harbor Dental | Spring newsletter design | Flat | $650.00 | 09/19/2026 | 09/24/2026 | Paid |
| Pine & Co. Realty | Listing photo edits (batch 3) | 6.5 hrs x $55.00 | $357.50 | | | Not invoiced |
| Lakeside Yoga | Website copy refresh | 3 hrs x $60.00 | $180.00 | | | Not invoiced |
| Main St. Gift Shop | Holiday flyer | Flat | $300.00 | | | Not invoiced |
| Harbor Dental | Social media templates | Flat | $480.00 | 09/05/2026 | | Waiting |

The social media templates invoice was 20 days old on 09/25/2026, so it shows Waiting. On 10/06/2026 it would switch to Overdue.

The Summary tab showed Open projects 3, Billed, not paid $480.00, and Paid $650.00. We checked these results in LibreOffice Calc 24.2.

![Projects tab of the freelance tracker with sample rows](images/2026-09-30-a-freelance-project-tracker-google-sheets/projects-tab-libreoffice.png) *The Projects tab with sample data, recalculated on 09/25/2026. PDF export from LibreOffice Calc 24.2; Days left and Payment change with today's date.*

## Troubleshooting

### Days left shows a date instead of a number
The cell picked up a date format. Select the Days left column and, in Google Sheets, click **Format > Number** and choose **Number**.

### The dropdowns are missing after import
The Status and Billing lists may not have come through. Add them again with **Data > Data validation** as shown in Build it yourself, using Hourly and Flat for Billing.

### Amount shows $0.00
Billing is blank or set to Hourly with Hours or Rate empty. A blank Billing makes the formula use Flat fee, so fill Billing first, then the matching fields.

### Days left or Payment shows #VALUE!
A date was stored as text, not a real date, often because the spreadsheet expects a different date order. `TODAY()` can't subtract text. In Google Sheets, check **File > Settings** and set Locale to United States, then retype the dates as MM/DD/YYYY. Google's help notes that the locale sets the spreadsheet's default date formatting.

### The numbers changed since yesterday
That's expected. Days left and Payment use `TODAY()`, which recalculates each time the file opens or changes.

## Template

[Download the freelance project tracker (.xlsx)](templates/tidy-tabs-freelance-project-tracker.xlsx)

The file has three tabs:

- **Projects:** 14 columns with the five sample rows, Status and Billing dropdowns through row 200, green shading for Done rows, and red for late Days left or Overdue payments.
- **Summary:** the four totals above.
- **How to use:** short instructions.

The sample clients are fictional. Delete rows 2 to 6 on the Projects tab to start fresh. To use it in Google Sheets, open a spreadsheet and go to **File > Import**, then upload the .xlsx.
