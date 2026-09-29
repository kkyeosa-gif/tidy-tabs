---
threads_url: https://www.threads.com/@tin_ylab/post/Dd2bXx8D2ut
blogger_url: https://tidytabs.blogspot.com/2026/09/how-to-build-client-contact-list.html
title: How to Build a Client Contact List Template in Google Sheets
labels: google-sheets-templates, crm
tested_in: LibreOffice Calc 24.2.7 (Linux), recalculated 09/28/2026. Google Sheets steps from Google Docs Editors Help (not tested here).
image_prompts: A freelancer's desk with a wooden index card box holding client contact cards, several tabbed with colored sticky flags, beside a phone and a desk calendar with a date circled in pen, card text blurred and unreadable.
image_alt: Index card box of client contacts with colored tabs flagging follow-up calls
search_description: Free Google Sheets template that flags clients due for a follow-up, using Last contact plus a follow-up interval and TODAY().
threads: One formula flags every client who's overdue: =IF(TODAY()>=E2,"Follow up now","OK").\nE2 is just Last contact plus your follow-up interval in days.\nNo more scanning a client list by eye to remember who you owe a call.
---
Track when you last talked to each client and flag who's due for a follow-up with one formula: `=IF(TODAY()>=E2,"Follow up now","OK")`, where E2 is Last contact plus your follow-up interval in days. Download the template below, or add the same three formula columns to a sheet you already use.

> Works in: Google Sheets (import steps from Google Docs Editors Help, not tested here). Tested in: LibreOffice Calc 24.2.7 (Linux), recalculated 09/28/2026.

![Client rows showing last contact dates, follow-up intervals, and a highlighted overdue status](images/2026-09-29-team-02-client-contact-list-template-google-sheets/client-contact-list-template.png) *LibreOffice Calc 24.2 PDF export of the Clients tab with sample data.*

## Steps

### In Google Sheets
1. Download the template from the Template section below.
2. Open a Google Sheets spreadsheet and click **File > Import**.
3. Choose **Upload** and select the .xlsx file from your computer.
4. Pick **Insert new sheet(s)** to add the Clients tab to a workbook you already use, or **Replace spreadsheet** to start fresh.
5. Click **Import data**. Google lists every import option on its [import help page](https://support.google.com/docs/answer/40608).

### In Excel
1. Download the template from the Template section below.
2. Open the .xlsx file directly. Excel reads the Clients tab, the formulas, and the red conditional formatting rule with no import step needed.

### Fill in your clients
1. On the Clients tab, replace the sample rows with your own: Client, Company, Last contact, and Follow up every (days).
2. Type Last contact as a real date, MM/DD/YYYY. If it lands lined up on the left side of the cell instead of the right, the spreadsheet stored it as text, and the formulas below won't work against it.
3. Leave Next follow-up and Status alone. Both are formulas.
4. Read Status: "Follow up now" means today's date has reached or passed Next follow-up. "OK" means you still have time.
5. After you reach out to a client, type today's date into Last contact. Next follow-up and Status update on their own.
6. Use Notes for anything you want on record, like what the last call was about.

### Build it yourself
If you'd rather add these columns to a sheet you already have, use these column letters: A Client, B Company, C Last contact, D Follow up every (days), E Next follow-up, F Status, G Notes.

Row 2 formulas, copied down:
- Next follow-up (E2): `=C2+D2`
- Status (F2): `=IF(E2="","",IF(TODAY()>=E2,"Follow up now","OK"))`

To shade "Follow up now" rows red in Google Sheets:
1. Select F2:F200, or however many rows you use.
2. Click **Format > Conditional formatting**.
3. Under **Format cells if**, choose **Custom formula is** and type `=$F2="Follow up now"`.
4. Pick a red fill and click **Done**.

Google's [conditional formatting help page](https://support.google.com/docs/answer/78413) covers custom formulas in more detail.

## Example

Here's how the four sample clients worked out when the formulas recalculated on 09/28/2026:

| Client | Last contact | Follow up every (days) | Next follow-up | Status |
| --- | --- | --- | --- | --- |
| Dana Ruiz | 09/01/2026 | 21 | 09/22/2026 | Follow up now |
| Emma Lee | 09/20/2026 | 14 | 10/04/2026 | OK |
| Sam Patel | 08/15/2026 | 30 | 09/14/2026 | Follow up now |
| Jamie Chen | 09/25/2026 | 7 | 10/02/2026 | OK |

Dana's and Sam's Next follow-up dates already passed by 09/28/2026, so both read Follow up now. Emma's and Jamie's are still ahead, so both read OK. These four clients and their dates are fictional sample data.

Because Status runs on `TODAY()`, these exact results only hold for 09/28/2026. Open the same file next week and some rows that show OK now will have flipped to Follow up now, and the reverse never happens until you update Last contact.

## Troubleshooting

### Status shows "OK" for a client who's clearly overdue
Last contact was typed or pasted in as text instead of a real date, so `TODAY()>=E2` compares a date to text and never comes back true. Delete the entry and retype the date directly into the cell as MM/DD/YYYY.

### Next follow-up shows a five-digit number instead of a date
The cell picked up a plain number format instead of a date format. Select the Next follow-up column and click **Format > Number > Date**.

### The red highlight didn't come through after import
Google Sheets import brings over values and formulas, but a conditional formatting rule can get dropped along the way. Select the Status column, click **Format > Conditional formatting**, and re-add the rule from Build it yourself above.

### Every row says "Follow up now"
Follow up every (days) is blank or set to 0, which makes Next follow-up equal to Last contact, so today is always past it. Fill in a real interval, like 14 or 30 days.

### The results changed since yesterday
That's expected. Status runs on `TODAY()`, so it recalculates every time the file opens or a cell changes, which is why the Example table above is only accurate for 09/28/2026.

## Template

[Download the .xlsx](templates/tidy-tabs-client-contact-list.xlsx)

The file has two tabs: **Clients**, with the four sample contacts and the Status conditional formatting already applied, and **How to use**, with short instructions. The sample clients are fictional; delete rows 2 through 5 on the Clients tab to start fresh with your own.

To open it in Google Sheets, go to **File > Import** and upload the .xlsx.
