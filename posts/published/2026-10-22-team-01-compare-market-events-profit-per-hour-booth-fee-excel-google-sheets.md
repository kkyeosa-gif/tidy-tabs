---
threads_url: https://www.threads.com/@tin_ylab/post/DezeYruFUNV
blogger_url: https://tidytabs.blogspot.com/2026/10/how-to-compare-farmers-markets-and.html
title: How to Compare Farmers Markets and Craft Fairs by Profit Per Hour in Excel and Google Sheets
labels: excel-formulas, small-business, sales
tested_in: LibreOffice Calc 24.2.7.2 (Linux) only. The template was recalculated there and compared with a Python Decimal model, including 17 edited variants of the events file (zero, blank and text hours, a misspelled event name, an event with no sales, tied events, negative profit, a text fee, an event listed twice). No mismatches in the final run. Excel and Google Sheets were not opened; the help links are from Microsoft Support and Google Docs Editors Help.
search_description: Compare farmers markets and craft fairs by profit and profit per hour in Excel and Google Sheets. Free event comparison template with sample data.
image_prompts: A wooden cash box with a few bills, a folded stack of canvas tote bags and a canopy weight bag on a folding market table at dawn, soft morning light, no screens, no readable text, no logos
image_alt: Cash box, tote bags and canopy weight on a folding table at an outdoor market
threads: Biggest market payout is not always your best market.\nA craft fair made $1,665.00 but took 18 hours, so it ranks 4th at $92.50 an hour.\nA 4 hour night market made $980.00, which is $245.00 an hour.
---
Find profit per hour for an event with `=IF(E2>0,F2/E2,"Check hours")`, where F2 is the event's profit (`=B2-C2-D2`, sales minus booth fee minus other costs) and E2 is the hours you spent there. The event with the most profit in dollars is not always the one with the most profit per hour.

> Works in: LibreOffice Calc 24.2.7.2 (Linux), where the template was built and every number below was recalculated and compared with a Python model. The Excel and Google Sheets steps use the same functions as typed. Neither Excel nor Google Sheets was opened for this post.

![Events sheet with sales, booth fees, profit per hour and rank for six events](images/2026-10-22-team-01-compare-market-events-profit-per-hour-booth-fee-excel-google-sheets/market-event-comparison-template.png) *caption: PDF export of the template with fictional sample data from LibreOffice Calc 24.2*

## Steps

Set up one row per event on an **Events** tab, and one row per sale on a **Sales** tab. The columns below match the template.

1. On the **Sales** tab, type one row per sale or shift total: **A** Date, **B** Event, **C** Amount, **D** Note (optional).
2. On the **Events** tab, type each event name once in **A2:A13**. Spell it exactly as it appears in Sales column B.
3. Type the booth fee in **C**, your other costs (gas, bags, supplies) in **D**, and the hours you spent at the event in **E**. Count setup and teardown if you want them in the rate.
4. In **B2**, add up that event's sales: `=SUMIFS(Sales!$C$2:$C$200,Sales!$B$2:$B$200,A2)`.
5. In **F2**, find the profit: `=B2-C2-D2`.
6. In **G2**, divide by hours: `=IF(E2>0,F2/E2,"Check hours")`.
7. Format B, C, D, F and G as currency with two decimals. Copy row 2 down for the other events.

The IF shows "Check hours" instead of an error when the hours are 0 or blank.

### What the template adds

The template cells are longer than the bare formulas above because they guard against blank, text and zero values:

- **B2** returns blank when A2 is blank, so empty rows stay empty. The SUMIFS inside is exactly step 4.
- **F2** shows "Check costs" if a booth fee or cost cell holds text such as "$35". A blank fee or cost counts as $0.00.
- **G2** tests that E2 is a number before it divides. This matters because Excel treats text as greater than 0, so a bare `E2>0` can pass text and then divide by it. The bare formula was not run on text for this post.
- **H2** is sales per hour and **I2** is the rank by profit per hour (1 is best, ties share a rank, "Check" rows are skipped).

The summary in **K1:L8** shows total sales, the best event by profit, the best event by profit per hour, and a CHECK cell in **L8**. It says OK when the event totals add up to the Sales tab and every amount is a number.

### In Excel

The formulas work as typed. Microsoft describes the arguments on its [SUMIFS function](https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b) page.

### In Google Sheets

Google Sheets has SUMIFS too, described on Google's [SUMIFS help page](https://support.google.com/docs/answer/3238496). The same formulas should work as typed, but Sheets was not opened for this post.

## Example

Sample data is fictional: six events for a small craft and food seller, from 09/12/2026 to 10/17/2026.

| Event | Sales | Booth fee | Other costs | Hours | Profit | Profit per hour | Rank |
|---|---|---|---|---|---|---|---|
| Riverside Farmers Market | $848.00 | $35.00 | $126.00 | 8 | $687.00 | $85.88 | 5 |
| Weekend Craft Fair | $2,310.00 | $175.00 | $470.00 | 18 | $1,665.00 | $92.50 | 4 |
| School PTA Fair | $620.00 | $40.00 | $85.00 | 5 | $495.00 | $99.00 | 3 |
| Sunday Farmers Market | $540.00 | $30.00 | $95.00 | 6 | $415.00 | $69.17 | 6 |
| Harvest Festival | $2,100.00 | $250.00 | $360.00 | 9 | $1,490.00 | $165.56 | 2 |
| Downtown Night Market | $1,180.00 | $60.00 | $140.00 | 4 | $980.00 | $245.00 | 1 |

Total sales are $7,598.00, and L8 shows OK.

The Weekend Craft Fair made the most profit, $1,665.00 over two days (10/10/2026 and 10/11/2026). But it took 18 hours, so it ranks fourth per hour. The Downtown Night Market made $980.00 in 4 hours, which is $245.00 an hour and first place.

This is arithmetic on fictional numbers, not advice on which events to apply to. Booth fees and costs here are just subtractions, not a statement about taxes.

## Troubleshooting

### An event's sales are less than I made
The name on the Sales tab does not match the name on the Events tab. A typo such as "Harvest Festvial" or a trailing space means SUMIFS skips those rows. L8 shows CHECK when this happens. Fix the spelling on either tab. Capital letters do not matter.

### The cell says Check hours
The hours cell is blank, 0, or text such as "8h". Type a plain number. Decimals work, so 7.5 is fine.

### The cell says Check costs
A booth fee or other costs cell holds text, such as "$35" typed as text. Retype it as a number.

### The best event by dollars is not the best per hour
That is expected when hours differ a lot. In the example the craft fair has the most profit and ranks fourth per hour because of its 18 hours. Compare the Profit and Profit per hour columns side by side.

### An event I listed twice is counted twice
If the same event name appears on two rows of the Events tab, its sales are added to both rows and the total is too high. L8 shows CHECK. Keep one row per event.

I did not test pasting event names from a web page, which can carry non-breaking spaces that look like normal spaces.

## Template

[Download the .xlsx](templates/tidy-tabs-market-event-comparison.xlsx)

The file has three tabs. **Sales** holds the yellow inputs in **A2:D200** (40 fictional sales rows are filled in). **Events** holds the inputs in **A2:A13** and **C2:E13** (up to 12 events) and the results in **B**, **F**, **G**, **H** and **I**, with the summary in **K1:L8**. **How to use** has short notes. All tabs print portrait on US Letter, and the summary beside the Events table does not print.

The data is fictional. The file was recalculated only in LibreOffice Calc 24.2.7.2, not in Excel or Google Sheets. To open it in Google Sheets, go to **File > Import > Upload** and select the file.
