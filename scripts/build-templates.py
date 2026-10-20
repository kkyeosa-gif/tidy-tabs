#!/usr/bin/env python3
"""Builds the downloadable .xlsx/.csv templates in templates/.

Run from the repo root: python3 scripts/build-templates.py
Needs openpyxl (pip install openpyxl). Every template is plain formulas
(IF, SUMIFS, TODAY, TEXT) so it opens the same in Excel and after
File > Import in Google Sheets. Re-run after editing; the files are
committed so readers download exactly what was checked.
"""
import csv
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path("templates")
USD = '"$"#,##0.00'
US_DATE = "mm/dd/yyyy"
US_DATETIME = "mm/dd/yyyy h:mm AM/PM"
HEADER_FILL = PatternFill("solid", fgColor="2E5E4E")
HEADER_FONT = Font(bold=True, color="FFFFFF")
ALERT_FILL = PatternFill("solid", fgColor="F8D7DA")
DONE_FILL = PatternFill("solid", fgColor="E2F0D9")
FORMULA_FILL = PatternFill("solid", fgColor="EDEDED")


def header(ws, names, widths):
    ws.append(names)
    for i, (cell, w) in enumerate(zip(ws[1], widths), start=1):
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"


def notes_sheet(wb, lines):
    ws = wb.create_sheet("How to use")
    ws.column_dimensions["A"].width = 100
    for line in lines:
        ws.append([line])
    ws["A1"].font = Font(bold=True, size=13)


def inventory():
    wb = Workbook()
    items = wb.active
    items.title = "Items"
    header(items, ["SKU", "Item", "Category", "Unit cost", "Sale price", "Starting qty",
                   "Qty in", "Qty out", "On hand", "Reorder at", "Status", "Stock value (cost)"],
           [10, 30, 14, 11, 11, 11, 9, 9, 9, 11, 13, 16])
    rows = [
        ("CND-LAV8", "Lavender soy candle, 8 oz", "Candles", 4.10, 18.00, 40, 10),
        ("CND-CED8", "Cedar soy candle, 8 oz", "Candles", 4.25, 18.00, 25, 10),
        ("SOP-OAT", "Oatmeal goat milk soap bar", "Soap", 1.85, 8.00, 60, 20),
        ("SOP-CHR", "Charcoal soap bar", "Soap", 2.05, 8.50, 30, 20),
        ("EAR-HOOP", "Brass hoop earrings", "Jewelry", 3.40, 24.00, 15, 5),
        ("CRD-BDAY", "Letterpress birthday card", "Cards", 0.95, 6.00, 80, 25),
        ("PKG-BOX6", "Kraft mailer box 6x6x4", "Supplies", 0.62, 0, 100, 40),
    ]
    for r, (sku, name, cat, cost, price, start, reorder) in enumerate(rows, start=2):
        items.append([sku, name, cat, cost, price, start,
                      f"=SUMIFS('Stock log'!$C:$C,'Stock log'!$B:$B,A{r})",
                      f"=SUMIFS('Stock log'!$D:$D,'Stock log'!$B:$B,A{r})",
                      f"=F{r}+G{r}-H{r}", reorder,
                      f'=IF(A{r}="","",IF(I{r}<=J{r},"REORDER","OK"))',
                      f"=I{r}*D{r}"])
        for col in "DEL":
            items[f"{col}{r}"].number_format = USD
        for col in "GHIKL":
            items[f"{col}{r}"].fill = FORMULA_FILL
    # Ranges run to row 500 (not just the sample rows) so a product added by
    # copying the last row down is counted, offered in the dropdown, and
    # highlighted without editing any formula. The total sits beside the
    # table, not under it, for the same reason.
    last = 500
    items.column_dimensions["N"].width = 18
    items.column_dimensions["O"].width = 14
    items["N1"], items["O1"] = "Total stock value", f"=SUM(L2:L{last})"
    items["N1"].font = Font(bold=True)
    items["O1"].number_format = USD
    items.conditional_formatting.add(f"K2:K{last}", FormulaRule(formula=[f'$K2="REORDER"'], fill=ALERT_FILL))

    log = wb.create_sheet("Stock log")
    header(log, ["Date", "SKU", "Qty in", "Qty out", "Note"], [12, 12, 9, 9, 40])
    entries = [
        (date(2026, 9, 1), "CND-LAV8", 24, 0, "Poured batch #14"),
        (date(2026, 9, 3), "CND-LAV8", 0, 12, "Etsy orders 9/1-9/3"),
        (date(2026, 9, 5), "SOP-OAT", 0, 18, "Saturday farmers market"),
        (date(2026, 9, 5), "CND-CED8", 0, 9, "Saturday farmers market"),
        (date(2026, 9, 8), "SOP-CHR", 0, 14, "Wholesale order, Main St. Gift Shop"),
        (date(2026, 9, 10), "EAR-HOOP", 0, 11, "Etsy orders"),
        (date(2026, 9, 12), "PKG-BOX6", 0, 64, "Shipped Etsy orders"),
        (date(2026, 9, 15), "CRD-BDAY", 0, 30, "Craft fair"),
    ]
    for e in entries:
        log.append(list(e))
        log.cell(row=log.max_row, column=1).number_format = US_DATE
    sku_list = DataValidation(type="list", formula1=f"=Items!$A$2:$A${last}", allow_blank=True)
    log.add_data_validation(sku_list)
    sku_list.add("B2:B1000")

    notes_sheet(wb, [
        "Tidy Tabs: Small Craft Business Inventory Tracker",
        "",
        "1. Items tab: one row per product or supply. Type your own SKU, cost, price, starting count, and reorder point.",
        "2. Stock log tab: add a row every time stock comes in (Qty in) or goes out (Qty out). Pick the SKU from the dropdown.",
        "3. On hand, Status, and Stock value update by themselves. Status turns red when On hand is at or below Reorder at.",
        "4. Don't type over the gray formula columns (Qty in, Qty out, On hand, Status, Stock value).",
        "5. To add a product, copy the last Items row down so the formulas come with it. Total stock value is in cell O1.",
        "",
        "Sample data is fictional. Delete the Stock log rows and change the Items rows to start fresh.",
        "Google Sheets: File > Import > Upload this .xlsx > Insert new sheet(s) or Replace spreadsheet.",
    ])
    wb.save(OUT / "tidy-tabs-craft-inventory-tracker.xlsx")


def projects():
    wb = Workbook()
    ws = wb.active
    ws.title = "Projects"
    header(ws, ["Client", "Project", "Start date", "Due date", "Status", "Billing", "Hours", "Rate",
                "Flat fee", "Amount", "Invoice sent", "Paid on", "Days left", "Payment"],
           [18, 28, 12, 12, 13, 10, 8, 9, 10, 12, 12, 12, 10, 12])
    rows = [
        ("Harbor Dental", "Spring newsletter design", date(2026, 9, 2), date(2026, 9, 19), "Done", "Flat", None, None, 650, date(2026, 9, 19), date(2026, 9, 24)),
        ("Pine & Co. Realty", "Listing photo edits (batch 3)", date(2026, 9, 8), date(2026, 9, 26), "In progress", "Hourly", 6.5, 55, None, None, None),
        ("Lakeside Yoga", "Website copy refresh", date(2026, 9, 15), date(2026, 10, 10), "In progress", "Hourly", 3, 60, None, None, None),
        ("Main St. Gift Shop", "Holiday flyer", date(2026, 9, 22), date(2026, 10, 3), "Not started", "Flat", None, None, 300, None, None),
        ("Harbor Dental", "Social media templates", date(2026, 8, 18), date(2026, 9, 5), "Done", "Flat", None, None, 480, date(2026, 9, 5), None),
    ]
    for r, (client, proj, start, due, status, billing, hours, rate, flat, sent, paid) in enumerate(rows, start=2):
        ws.append([client, proj, start, due, status, billing, hours, rate, flat,
                   f'=IF(F{r}="Hourly",G{r}*H{r},I{r})', sent, paid,
                   f'=IF(OR(D{r}="",E{r}="Done"),"",D{r}-TODAY())',
                   f'=IF(L{r}<>"","Paid",IF(K{r}="","Not invoiced",IF(TODAY()-K{r}>30,"Overdue","Waiting")))'])
        for col in "CDKL":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "HIJ":
            ws[f"{col}{r}"].number_format = USD
    last = 200
    status = DataValidation(type="list", formula1='"Not started,In progress,Waiting on client,Done"', allow_blank=True)
    billing = DataValidation(type="list", formula1='"Hourly,Flat"', allow_blank=True)
    ws.add_data_validation(status)
    ws.add_data_validation(billing)
    status.add(f"E2:E{last}")
    billing.add(f"F2:F{last}")
    ws.conditional_formatting.add(f"A2:N{last}", FormulaRule(formula=['$E2="Done"'], fill=DONE_FILL))
    ws.conditional_formatting.add(f"M2:M{last}", FormulaRule(formula=['AND(ISNUMBER($M2),$M2<0)'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"N2:N{last}", FormulaRule(formula=['$N2="Overdue"'], fill=ALERT_FILL))

    summary = wb.create_sheet("Summary")
    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 14
    for label, formula in [
        ("Open projects", '=COUNTIFS(Projects!A2:A200,"<>",Projects!E2:E200,"<>Done")'),
        ("Billed, not paid ($)", '=SUMIFS(Projects!J2:J200,Projects!K2:K200,"<>",Projects!L2:L200,"")'),
        ("Paid ($)", '=SUMIFS(Projects!J2:J200,Projects!L2:L200,"<>")'),
        ("Invoices overdue (30+ days)", '=COUNTIF(Projects!N2:N200,"Overdue")'),
    ]:
        summary.append([label, formula])
    for r in (2, 3):
        summary[f"B{r}"].number_format = USD

    notes_sheet(wb, [
        "Tidy Tabs: Freelance Project Tracker",
        "",
        "1. One row per project. Pick Status and Billing from the dropdowns.",
        "2. Hourly jobs: fill Hours and Rate. Flat-fee jobs: fill Flat fee. Amount picks the right one.",
        "3. Type the date you sent the invoice in Invoice sent, and the date the money arrived in Paid on.",
        "4. Days left counts down to the due date (negative = late) and hides itself once Status is Done.",
        "5. Payment shows Not invoiced / Waiting / Overdue (more than 30 days after you invoiced) / Paid.",
        "6. Days left and Payment use TODAY(), so they change every day you open the file.",
        "",
        "Sample clients are fictional. Delete rows 2-6 on the Projects tab to start fresh.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-freelance-project-tracker.xlsx")


ZIPS = [
    ("Boston", "MA", "02108", "02108-1522"),
    ("Hoboken", "NJ", "07030", "07030-3707"),
    ("Holtsville", "NY", "00501", "00501-0001"),
    ("Burlington", "VT", "05401", "05401-4417"),
    ("Chicago", "IL", "60601", "60601-3014"),
    ("Las Vegas", "NV", "89101", "89101-5019"),
]


def zip_codes():
    wb = Workbook()
    ws = wb.active
    ws.title = "Fixed (text)"
    header(ws, ["City", "State", "ZIP (text)", "ZIP+4 (text)"], [14, 7, 12, 14])
    for city, st, z, z4 in ZIPS:
        ws.append([city, st, z, z4])
    for row in ws.iter_rows(min_row=2, max_row=1000, min_col=3, max_col=4):
        for c in row:
            c.number_format = "@"

    broken = wb.create_sheet("Broken (numbers)")
    header(broken, ["City", "State", "ZIP as typed", "Repair with TEXT", "Repair ZIP+4 (9 digits)"], [14, 7, 14, 18, 22])
    for r, (city, st, z, z4) in enumerate(ZIPS, start=2):
        broken.append([city, st, int(z), f'=TEXT(C{r},"00000")', f'=TEXT({int(z4.replace("-", ""))},"00000-0000")'])

    notes_sheet(wb, [
        "Tidy Tabs: US ZIP Code Leading Zeros sample",
        "",
        "Fixed (text): the ZIP columns are formatted as Text (@) BEFORE the codes were entered, so 02108 stays 02108.",
        "Broken (numbers): the same ZIPs stored as numbers, which is what happens when you type or paste them into a General cell.",
        'Column D rebuilds a 5-digit ZIP with =TEXT(C2,"00000"). Copy it, then Paste Special > Values to keep the result.',
        "Column E shows the ZIP+4 version with the format 00000-0000.",
        "",
        "The 5-digit ZIPs are real; the +4 extensions are made up for practice.",
        "The same six ZIP codes are in tidy-tabs-zip-codes-sample.csv for practicing a CSV import.",
    ])
    wb.save(OUT / "tidy-tabs-zip-codes-leading-zeros.xlsx")

    with open(OUT / "tidy-tabs-zip-codes-sample.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["City", "State", "ZIP", "ZIP+4"])
        for row in ZIPS:
            w.writerow(row)


def print_one_page():
    wb = Workbook()
    ws = wb.active
    ws.title = "Price list"
    cols = ["SKU", "Item", "Category", "Size", "Color", "Unit cost", "Wholesale", "Retail",
            "Margin %", "On hand", "Reorder at", "Supplier", "Lead time (days)", "Last ordered"]
    header(ws, cols, [10, 30, 12, 10, 10, 10, 11, 10, 10, 9, 12, 14, 10, 13])
    colors = ["Natural", "Sage", "Navy", "Rust", "Cream"]
    sizes = ["4 oz", "8 oz", "12 oz"]
    for i in range(1, 49):
        r = i + 1
        cost = round(1.5 + (i % 7) * 0.65, 2)
        ws.append([f"ITM-{i:03d}", f"Sample product {i}", ["Candles", "Soap", "Cards", "Jewelry"][i % 4],
                   sizes[i % 3], colors[i % 5], cost, f"=ROUND(F{r}*2,2)", f"=ROUND(F{r}*{3 + (i % 3) * 0.5},2)",
                   f"=(H{r}-F{r})/H{r}", 10 + (i * 7) % 50, 15, f"Supplier {chr(65 + i % 6)}",
                   7 + (i % 4) * 7, date(2026, 8, 1 + i % 28)])
        for col in "FGH":
            ws[f"{col}{r}"].number_format = USD
        ws[f"I{r}"].number_format = "0%"
        ws[f"N{r}"].number_format = US_DATE
    # The settings the post walks through, already applied: US Letter,
    # landscape, fit all columns on one page wide, header row repeats.
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_title_rows = "1:1"
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_options.horizontalCentered = True
    wb.save(OUT / "tidy-tabs-print-one-page-us-letter.xlsx")


ADDRESSES = [
    ("Boston, MA 02108", "Harbor Dental"),
    ("Hoboken, NJ 07030", "Pine & Co. Realty"),
    ("Holtsville, NY 00501", "Main St. Gift Shop"),
    ("Burlington, VT 05401", "Lakeside Yoga"),
    ("Chicago, IL 60601", "Northside Print Co."),
    ("Las Vegas, NV 89101", "Desert Bloom Candles"),
    ("Salt Lake City, UT 84101", "Wasatch Bike Repair"),
    ("St. Louis, MO 63101", "Gateway Bakery"),
]


def split_city_state_zip():
    wb = Workbook()
    ws = wb.active
    ws.title = "Split"
    header(ws, ["Customer", "City, ST ZIP", "City", "State", "ZIP"], [22, 28, 18, 8, 9])
    for r, (addr, name) in enumerate(ADDRESSES, start=2):
        ws.append([name, addr,
                   f'=LEFT(B{r},FIND(",",B{r})-1)',
                   f'=MID(B{r},FIND(",",B{r})+2,2)',
                   f'=RIGHT(B{r},5)'])
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    notes_sheet(wb, [
        "Tidy Tabs: Split City, ST ZIP into three columns",
        "",
        'Type or paste addresses like "Boston, MA 02108" in column B.',
        "City, State, and ZIP (gray) fill in by formula. Copy a gray row down for more addresses.",
        'City = LEFT(B2,FIND(",",B2)-1): everything before the comma.',
        'State = MID(B2,FIND(",",B2)+2,2): the 2 letters after the comma and space.',
        "ZIP = RIGHT(B2,5): the last 5 characters, returned as text, so 02108 keeps its 0.",
        "Works for 5-digit ZIPs. For ZIP+4 (02108-1522), use RIGHT(B2,10).",
        "To keep only the results, copy C:E and Paste Special > Values.",
        "",
        "Customer names are fictional; the city/state/ZIP combinations are real.",
    ])
    wb.save(OUT / "tidy-tabs-split-city-state-zip.xlsx")


def dependent_dropdown():
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    ws.title = "Order form"
    header(ws, ["Category", "Subcategory", "Item note", "Test: COUNTA(INDIRECT(cat))"],
           [14, 22, 30, 30])
    # Sample rows exercise every category so the INDIRECT lookup is checked
    # against each named range, not just one.
    rows = ["Candles", "Soap", "Jewelry", "Cards"]
    for r, cat in enumerate(rows, start=2):
        ws.append([cat, "", "pick from Subcategory dropdown", f"=COUNTA(INDIRECT(A{r}))"])
        ws[f"D{r}"].fill = FORMULA_FILL

    lists = wb.create_sheet("Lists")
    header(lists, ["Candles", "Soap", "Jewelry", "Cards"], [16, 16, 16, 16])
    cols = {
        "A": ["Lavender soy 8 oz", "Cedar soy 8 oz", "Vanilla soy 4 oz"],
        "B": ["Oatmeal goat milk bar", "Charcoal bar"],
        "C": ["Brass hoop earrings", "Beaded necklace", "Cuff bracelet", "Stacking ring"],
        "D": ["Letterpress birthday card", "Letterpress thank you card"],
    }
    for col, items in cols.items():
        for i, val in enumerate(items, start=2):
            lists[f"{col}{i}"] = val

    # Named ranges sized exactly to each list, keyed by the category text so
    # =INDIRECT(A2) resolves "Candles" to the Candles named range.
    ranges = {"Candles": "A2:A4", "Soap": "B2:B3", "Jewelry": "C2:C5", "Cards": "D2:D3"}
    for name, rng in ranges.items():
        wb.defined_names[name] = DefinedName(name, attr_text=f"Lists!${rng.split(':')[0][0]}${rng.split(':')[0][1:]}:${rng.split(':')[1][0]}${rng.split(':')[1][1:]}")

    cat_list = DataValidation(type="list", formula1="=Lists!$A$1:$D$1", allow_blank=True)
    ws.add_data_validation(cat_list)
    cat_list.add("A2:A200")
    sub_list = DataValidation(type="list", formula1="=INDIRECT($A2)", allow_blank=True)
    ws.add_data_validation(sub_list)
    sub_list.add("B2:B200")

    notes_sheet(wb, [
        "Tidy Tabs: Dependent (Category > Subcategory) Drop-Down List",
        "",
        "Order form tab: pick a Category in column A, then Subcategory in column B only shows items for that category.",
        "Lists tab: each column header is a category name; the cells below it are that category's subcategories.",
        "Behind the scenes, each Lists column is a named range with the same name as its header (Candles, Soap, Jewelry, Cards).",
        "Subcategory's data validation formula is =INDIRECT($A2), which turns the Category text into the matching named range.",
        "Column D is a check, not part of the form: =COUNTA(INDIRECT(A2)) counts how many subcategories INDIRECT found for that row's category, so you can confirm the link works before you rely on it.",
        "To add a category: add a column to Lists, name a range after it (Data > Named ranges), and add that name to the Category dropdown list.",
        "",
        "Sample categories and products are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx. Named ranges and INDIRECT dropdowns import with it.",
    ])
    wb.save(OUT / "tidy-tabs-dependent-dropdown-list.xlsx")


def client_contact_list():
    wb = Workbook()
    ws = wb.active
    ws.title = "Clients"
    header(ws, ["Client", "Company", "Last contact", "Follow up every (days)",
                "Next follow-up", "Status", "Notes"], [16, 22, 13, 15, 13, 14, 30])
    rows = [
        ("Dana Ruiz", "Harbor Dental", date(2026, 9, 1), 21, "Wants a spring newsletter quote"),
        ("Emma Lee", "Pine & Co. Realty", date(2026, 9, 20), 14, "Sent listing photo proofs"),
        ("Sam Patel", "Lakeside Yoga", date(2026, 8, 15), 30, "Website copy in review"),
        ("Jamie Chen", "Main St. Gift Shop", date(2026, 9, 25), 7, "Holiday flyer kickoff call"),
    ]
    for r, (name, company, last, every, notes) in enumerate(rows, start=2):
        ws.append([name, company, last, every, f"=C{r}+D{r}",
                   f'=IF(E{r}="","",IF(TODAY()>=E{r},"Follow up now","OK"))', notes])
        for col in "CE":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "EF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    last = 300
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['$F2="Follow up now"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Client Contact List with Follow-Up Reminders",
        "",
        "1. One row per client. Fill in Last contact and Follow up every (days).",
        "2. Next follow-up = Last contact + Follow up every (days), calculated for you.",
        "3. Status turns red and reads \"Follow up now\" once today's date reaches Next follow-up.",
        "4. Status uses TODAY(), so it updates on its own every day you open the file.",
        "5. After you reach out, update Last contact to today's date to reset the countdown.",
        "",
        "Sample clients are fictional. Delete rows 2-5 to start your own list.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-client-contact-list.xlsx")


def packing_slip():
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing slip"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws["A1"] = "Tidy Tabs Craft Co."
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "PACKING SLIP"
    ws["A2"].font = Font(bold=True, size=12)
    ws["A4"] = "Order #"
    ws["B4"] = "1042"
    ws["A5"] = "Order date"
    ws["B5"] = date(2026, 9, 26)
    ws["B5"].number_format = US_DATE
    ws["A6"] = "Ship to"
    ws["B6"] = "Priya Nair"
    ws["B7"] = "48 Willow St"
    ws["B8"] = "Burlington, VT 05401"

    headers_row = 10
    for col, val in zip("ABCD", ["Item", "SKU", "Qty ordered", "Qty packed"]):
        c = ws[f"{col}{headers_row}"]
        c.value = val
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    items = [
        ("Lavender soy candle, 8 oz", "CND-LAV8", 2, 2),
        ("Charcoal soap bar", "SOP-CHR", 3, 3),
        ("Brass hoop earrings", "EAR-HOOP", 1, 1),
        ("Letterpress birthday card", "CRD-BDAY", 4, 3),
    ]
    for i, (name, sku, qty_o, qty_p) in enumerate(items):
        r = headers_row + 1 + i
        ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = name, sku, qty_o, qty_p
    last_item_row = headers_row + len(items)
    total_row = last_item_row + 2
    ws[f"C{total_row}"] = "Total packed"
    ws[f"C{total_row}"].font = Font(bold=True)
    ws[f"D{total_row}"] = f"=SUM(D{headers_row + 1}:D{last_item_row})"
    ws[f"D{total_row}"].font = Font(bold=True)
    ws[f"A{total_row + 2}"] = "Thank you for your order! Questions: hello@example.com"

    # US Letter, portrait, one page, matching the print-one-page settings
    # used on the other printable templates in this file.
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = ws.page_margins.right = 0.75
    ws.page_margins.top = ws.page_margins.bottom = 0.75

    notes_sheet(wb, [
        "Tidy Tabs: Packing Slip Template",
        "",
        "1. Fill in Order #, Order date, and Ship to at the top.",
        "2. List one item per row: Item, SKU, Qty ordered, Qty packed.",
        "3. Total packed adds up column D on its own: =SUM(D11:D14) in this sample.",
        "4. If Qty packed is less than Qty ordered for any row, that item shipped short; note why in a blank row below the table.",
        "5. Page is already set to US Letter, portrait, and fit to one page, so File > Print needs no changes.",
        "",
        "Business name, order, and customer are fictional sample data.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-packing-slip-template.xlsx")


def running_balance():
    wb = Workbook()
    ws = wb.active
    ws.title = "Cash log"
    header(ws, ["Date", "Description", "Money in", "Money out", "Balance"], [12, 32, 11, 11, 12])
    ws.append([date(2026, 9, 1), "Opening balance", None, None, 250.00])
    entries = [
        (date(2026, 9, 6), "Farmers market cash sales", 180.00, None),
        (date(2026, 9, 9), "Bought jars and labels", None, 45.50),
        (date(2026, 9, 15), "Etsy payout deposited", 96.20, None),
        (date(2026, 9, 20), "Paid booth fee, Oct market", None, 35.00),
    ]
    for desc_row, (d, desc, in_amt, out_amt) in enumerate(entries, start=3):
        prev = desc_row - 1
        ws.append([d, desc, in_amt, out_amt, f"=E{prev}+C{desc_row}-D{desc_row}"])
    for r in range(2, 2 + 1 + len(entries)):
        ws[f"A{r}"].number_format = US_DATE
        for col in "CDE":
            ws[f"{col}{r}"].number_format = USD
        ws[f"E{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Running Balance Cash Log",
        "",
        "1. Row 2 is your opening balance, typed in directly (not a formula).",
        "2. Every row after that: type Money in or Money out (leave the other blank).",
        "3. Balance adds the row above's balance, plus Money in, minus Money out: =E2+C3-D3, copied down.",
        "4. Never type over the Balance column (gray); copy the formula down when you add a row.",
        "",
        "Sample cash entries are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-running-balance-cash-log.xlsx")


def highlight_duplicates():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order #", "Customer email", "Duplicate?"], [12, 30, 14])
    rows = [
        ("1001", "priya.nair@example.com"),
        ("1002", "sam.patel@example.com"),
        ("1003", "dana.ruiz@example.com"),
        ("1004", "priya.nair@example.com"),  # repeat customer, different order
        ("1005", "emma.lee@example.com"),
        ("1006", "jamie.chen@example.com"),
        ("1007", "sam.patel@example.com"),   # repeat customer
        ("1008", "new.customer@example.com"),
        ("1009", "dana.ruiz@example.com"),   # repeat customer
        ("1010", "ana.gomez@example.com"),
        ("1011", "ana.gomez@example.com"),   # same order pasted in twice by mistake
        ("1012", "priya.nair@example.com"),  # third order, same email
        ("1013", "tom.reyes@example.com"),
        ("1014", "lee.wu@example.com"),
        ("1015", "sam.patel@example.com"),   # repeat customer
    ]
    last = 1 + len(rows)
    for r, (order, email) in enumerate(rows, start=2):
        ws.append([order, email, f'=IF(COUNTIF($B$2:$B${last},B{r})>1,"Duplicate","")'])
        ws[f"C{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"B2:B{last}", FormulaRule(formula=[f'COUNTIF($B$2:$B${last},B2)>1'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Highlight Duplicate Values Sample",
        "",
        "Column C shows =IF(COUNTIF($B$2:$B$16,B2)>1,\"Duplicate\",\"\") for every row.",
        "The same COUNTIF formula drives the red conditional formatting fill on column B (Format > Conditional formatting > Custom formula is).",
        "This flags repeated values; it never deletes or moves rows, so repeat customers with more than one real order still show up highlighted on purpose.",
        "To find only accidental duplicates (the exact same order entered twice), also compare the Order # column, or sort by email and scan by eye.",
        "",
        "Order numbers and emails are fictional sample data.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-highlight-duplicates-sample.xlsx")


def business_days_ship_by():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order #", "Customer", "Order date", "Business days to ship", "Ship by",
                "Calendar days", "Business days check"], [10, 22, 12, 14, 12, 12, 16])
    rows = [
        ("2001", "Harbor Dental", date(2026, 9, 28), 5),
        ("2002", "Lakeside Yoga", date(2026, 10, 6), 10),
        ("2003", "Main St. Gift Shop", date(2026, 11, 9), 3),
        ("2004", "Gateway Bakery", date(2026, 11, 20), 5),
    ]
    for r, (num, cust, d, n) in enumerate(rows, start=2):
        ws.append([num, cust, d, n,
                   f"=WORKDAY(C{r},D{r},Holidays!$A$2:$A$20)",
                   f"=E{r}-C{r}",
                   f"=NETWORKDAYS(C{r},E{r},Holidays!$A$2:$A$20)-1"])
        for col in "CE":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "EFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    hol = wb.create_sheet("Holidays")
    header(hol, ["Date", "Day off"], [12, 28])
    for d, name in [(date(2026, 9, 7), "Labor Day"), (date(2026, 10, 12), "Columbus / Indigenous Peoples' Day"),
                    (date(2026, 11, 11), "Veterans Day"), (date(2026, 11, 26), "Thanksgiving Day"),
                    (date(2026, 12, 25), "Christmas Day")]:
        hol.append([d, name])
        hol.cell(row=hol.max_row, column=1).number_format = US_DATE
    hol.column_dimensions["B"].width = 34

    notes_sheet(wb, [
        "Tidy Tabs: Business Days and Ship-By Dates",
        "",
        "1. Orders tab: type the Order date and how many business days you need to ship.",
        "2. Ship by = WORKDAY(C2,D2,Holidays!$A$2:$A$20). It skips Saturdays, Sundays, and every date on the Holidays tab.",
        "3. Calendar days shows how many real days that is. Business days check = NETWORKDAYS(C2,E2,Holidays!$A$2:$A$20)-1 counts back and should equal column D.",
        "4. Holidays tab: add or remove your own days off (market days, vacation). Ranges run to row 20.",
        "5. WORKDAY and NETWORKDAYS treat Saturday and Sunday as the weekend. For other weekends use WORKDAY.INTL.",
        "",
        "Customers and orders are fictional. Holiday dates are the 2026 US federal holidays that fall in the sample range; check your own calendar.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-business-days-ship-by-date.xlsx")


def time_to_decimal_hours():
    from datetime import time

    wb = Workbook()
    ws = wb.active
    ws.title = "Shifts"
    header(ws, ["Date", "Clock in", "Clock out", "Unpaid break (min)", "Time worked (h:mm)",
                "Decimal hours", "Nearest 15 min", "Pay"], [12, 11, 11, 13, 14, 12, 13, 11])
    rows = [
        (date(2026, 9, 21), time(8, 30), time(17, 0), 30),
        (date(2026, 9, 22), time(9, 0), time(17, 45), 30),
        (date(2026, 9, 23), time(7, 15), time(15, 40), 30),
        (date(2026, 9, 24), time(22, 0), time(6, 30), 30),
        (date(2026, 9, 25), time(9, 5), time(13, 20), 0),
    ]
    for r, (d, tin, tout, brk) in enumerate(rows, start=2):
        ws.append([d, tin, tout, brk,
                   f"=MOD(C{r}-B{r},1)-D{r}/1440",
                   f"=ROUND(E{r}*24,2)",
                   f"=MROUND(E{r}*24,0.25)",
                   f"=ROUND(F{r}*$K$1,2)"])
        ws[f"A{r}"].number_format = US_DATE
        for col in "BC":
            ws[f"{col}{r}"].number_format = "h:mm AM/PM"
        ws[f"E{r}"].number_format = "h:mm"
        ws[f"F{r}"].number_format = "0.00"
        ws[f"G{r}"].number_format = "0.00"
        ws[f"H{r}"].number_format = USD
        for col in "EFGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["J1"], ws["K1"] = "Hourly rate", 22.5
    ws["J2"], ws["K2"] = "Total decimal hours", "=SUM(F2:F200)"
    ws["J3"], ws["K3"] = "Total pay", "=SUM(H2:H200)"
    ws["K1"].number_format = USD
    ws["K3"].number_format = USD
    for c in ("J1", "J2", "J3"):
        ws[c].font = Font(bold=True)
    for col, w in (("J", 20), ("K", 12)):
        ws.column_dimensions[col].width = w

    notes_sheet(wb, [
        "Tidy Tabs: Convert Clock Times to Decimal Hours",
        "",
        "1. Type Clock in, Clock out, and Unpaid break in minutes. Format the time cells as Time.",
        "2. Time worked = MOD(C2-B2,1)-D2/1440. MOD makes overnight shifts (10:00 PM to 6:30 AM) come out positive; 1440 is minutes in a day.",
        "3. Decimal hours = ROUND(E2*24,2). Excel and Sheets store time as a fraction of a day, so multiply by 24 to get hours.",
        "4. Nearest 15 min = MROUND(E2*24,0.25), for payroll that rounds to quarter hours.",
        "5. Pay = ROUND(Decimal hours x the hourly rate in K1,2). Change the rate there.",
        "",
        "Sample shifts and the $22.50 rate are fictional. Rounding rules are set by your own payroll policy; this file only does the arithmetic.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-time-to-decimal-hours.xlsx")


def receivables_aging():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Invoice #", "Client", "Invoice date", "Terms (days)", "Due date", "Amount",
                "Paid on", "Balance", "Days past due", "Bucket"],
           [10, 22, 12, 11, 12, 12, 12, 12, 11, 11])
    rows = [
        ("1001", "Harbor Dental", date(2026, 7, 1), 30, 650, date(2026, 8, 5)),
        ("1002", "Pine & Co. Realty", date(2026, 7, 15), 30, 1200, None),
        ("1003", "Lakeside Yoga", date(2026, 8, 10), 30, 380, None),
        ("1004", "Main St. Gift Shop", date(2026, 8, 25), 15, 300, None),
        ("1005", "Northside Print Co.", date(2026, 9, 5), 30, 875, None),
        ("1006", "Desert Bloom Candles", date(2026, 5, 20), 30, 450, None),
        ("1007", "Wasatch Bike Repair", date(2026, 6, 30), 30, 220, None),
        ("1008", "Gateway Bakery", date(2026, 9, 20), 30, 540, None),
        ("1009", "Harbor Dental", date(2026, 9, 15), 30, 650, None),
    ]
    for r, (num, client, d, terms, amt, paid) in enumerate(rows, start=2):
        ws.append([num, client, d, terms, f"=C{r}+D{r}", amt, paid,
                   f'=IF(OR(F{r}="",G{r}<>""),0,F{r})',
                   f'=IF(H{r}=0,"",MAX(0,Aging!$B$1-E{r}))',
                   f'=IF(H{r}=0,"Paid",IF(I{r}=0,"Current",IF(I{r}<=30,"1-30",IF(I{r}<=60,"31-60",IF(I{r}<=90,"61-90","90+")))))'])
        for col in "CEG":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "FH":
            ws[f"{col}{r}"].number_format = USD
        for col in "EHIJ":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    ag = wb.create_sheet("Aging")
    ag.column_dimensions["A"].width = 22
    ag.column_dimensions["B"].width = 14
    ag.column_dimensions["C"].width = 12
    ag["A1"], ag["B1"] = "As of date", date(2026, 9, 30)
    ag["B1"].number_format = US_DATE
    ag["A1"].font = Font(bold=True)
    ag.append([])
    ag.append(["Bucket (days past due)", "Balance", "Invoices"])
    for c in ag[3]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for r, b in enumerate(["Current", "1-30", "31-60", "61-90", "90+"], start=4):
        ag.append([b, f"=SUMIFS(Invoices!$H$2:$H$500,Invoices!$J$2:$J$500,A{r})",
                   f"=COUNTIFS(Invoices!$J$2:$J$500,A{r})"])
        ag[f"B{r}"].number_format = USD
    ag.append(["Total outstanding", "=SUM(B4:B8)", "=SUM(C4:C8)"])
    ag["B9"].number_format = USD
    ag["A9"].font = ag["B9"].font = Font(bold=True)
    ws.conditional_formatting.add("J2:J500", FormulaRule(formula=['OR($J2="61-90",$J2="90+")'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Accounts Receivable Aging Report",
        "",
        "1. Invoices tab: one row per invoice. Fill Invoice #, Client, Invoice date, Terms (days), Amount. Type Paid on when the money arrives.",
        "2. Due date = Invoice date + Terms. Balance is the Amount until you enter a Paid on date, then 0.",
        "3. Days past due = As of date minus Due date (0 if not late yet). The As of date is cell B1 on the Aging tab.",
        "4. Bucket sorts each unpaid invoice into Current, 1-30, 31-60, 61-90, or 90+ days past due.",
        "5. Aging tab adds up Balance by bucket with SUMIFS. Type =TODAY() in B1 to always age as of today; the sample uses 09/30/2026 so the numbers stay put.",
        "",
        "Clients and amounts are fictional. This only groups invoices by how late they are; it makes no decision about collection or write-offs.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-receivables-aging-report.xlsx")


def markup_vs_margin():
    wb = Workbook()
    ws = wb.active
    ws.title = "Pricing"
    header(ws, ["Item", "Unit cost", "Markup %", "Price from markup", "Profit per unit", "Actual margin %",
                "Target margin %", "Price for target margin", "Margin check"],
           [26, 11, 11, 14, 13, 13, 13, 16, 13])
    rows = [
        ("Lavender soy candle, 8 oz", 4.10, 1.00, 0.50),
        ("Oatmeal goat milk soap bar", 1.80, 1.50, 0.60),
        ("Brass hoop earrings", 3.40, 2.00, 0.70),
        ("Letterpress birthday card", 0.95, 3.00, 0.75),
    ]
    for r, (item, cost, mk, tgt) in enumerate(rows, start=2):
        ws.append([item, cost, mk, f"=ROUND(B{r}*(1+C{r}),2)", f"=D{r}-B{r}", f"=E{r}/D{r}",
                   tgt, f"=ROUND(B{r}/(1-G{r}),2)", f"=(H{r}-B{r})/H{r}"])
        for col in "BDEH":
            ws[f"{col}{r}"].number_format = USD
        for col in "CFGI":
            ws[f"{col}{r}"].number_format = "0.0%"
        for col in "DEFHI":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Markup vs Profit Margin Calculator",
        "",
        "1. Type Unit cost and the Markup % you use (100% means price is double the cost).",
        "2. Price from markup = ROUND(B2*(1+C2),2). Profit per unit = price minus cost.",
        "3. Actual margin % = Profit / Price. A 100% markup is only a 50% margin, because margin is measured against the price, not the cost.",
        "4. Going the other way: type a Target margin %. Price for target margin = ROUND(B2/(1-G2),2). Do not use B2*(1+G2); that mixes up markup and margin.",
        "5. Margin check divides the profit at the rounded price by that price; it can differ from the target by a fraction of a percent.",
        "",
        "Items and costs are fictional. Prices are arithmetic only; they do not include shipping, fees, or taxes.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-markup-vs-margin-calculator.xlsx")


def vlookup_price_list():
    wb = Workbook()
    ws = wb.active
    ws.title = "Order lines"
    header(ws, ["SKU", "Qty", "Item", "Unit price", "Line total", "Price via INDEX/MATCH"],
           [12, 7, 30, 11, 12, 20])
    rows = [("CND-LAV8", 2), ("SOP-OAT", 3), ("EAR-HOOP", 1), ("SOP-OAT ", 1),
            ("CRD-BDAY", 5), ("CND-CED8", 1)]
    for r, (sku, qty) in enumerate(rows, start=2):
        ws.append([sku, qty,
                   f'=IFERROR(VLOOKUP(A{r},\'Price list\'!$A$2:$C$500,2,FALSE),"SKU not found")',
                   f'=IFERROR(VLOOKUP(A{r},\'Price list\'!$A$2:$C$500,3,FALSE),"")',
                   f'=IF(D{r}="","",B{r}*D{r})',
                   f'=IFERROR(INDEX(\'Price list\'!$C$2:$C$500,MATCH(A{r},\'Price list\'!$A$2:$A$500,0)),"")'])
        for col in "DEF":
            ws[f"{col}{r}"].number_format = USD
        for col in "CDEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["H1"], ws["I1"] = "Order total", "=SUM(E2:E500)"
    ws["H1"].font = Font(bold=True)
    ws["I1"].number_format = USD
    ws.column_dimensions["H"].width = 14
    ws.column_dimensions["I"].width = 12

    pl = wb.create_sheet("Price list")
    header(pl, ["SKU", "Item", "Price"], [12, 30, 11])
    for sku, item, price in [
        ("CND-LAV8", "Lavender soy candle, 8 oz", 18.00), ("CND-CED8", "Cedar soy candle, 8 oz", 18.00),
        ("SOP-OAT", "Oatmeal goat milk soap bar", 8.00), ("SOP-CHR", "Charcoal soap bar", 8.50),
        ("EAR-HOOP", "Brass hoop earrings", 24.00), ("CRD-BDAY", "Letterpress birthday card", 6.00),
    ]:
        pl.append([sku, item, price])
        pl.cell(row=pl.max_row, column=3).number_format = USD

    notes_sheet(wb, [
        "Tidy Tabs: Look Up a Price with VLOOKUP",
        "",
        "1. Price list tab: SKU must be the first column, then Item, then Price.",
        "2. Order lines tab: type a SKU and Qty. Item = VLOOKUP(A2,'Price list'!$A$2:$C$500,2,FALSE). FALSE means exact match only.",
        "3. Unit price uses column number 3. Line total = Qty x Unit price. Order total is in I1.",
        "4. IFERROR turns a missing SKU into \"SKU not found\" instead of #N/A. Row 5 has a trailing space after SOP-OAT on purpose, so it is not found.",
        "5. Column F does the same lookup with INDEX/MATCH, which also works when the lookup column is not first.",
        "",
        "Products and prices are fictional. XLOOKUP is not used here because it needs newer versions of Excel and LibreOffice.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-vlookup-price-list.xlsx")


def break_even_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Break-even"
    header(ws, ["Product", "Price", "Variable cost per unit", "Fixed costs", "Profit per unit",
                "Break-even units", "Break-even revenue"], [28, 10, 14, 12, 12, 13, 15])
    rows = [
        ("Lavender soy candle, 8 oz", 18.00, 6.50, 120.00),
        ("Oatmeal goat milk soap bar", 8.00, 3.25, 75.00),
        ("Brass hoop earrings", 24.00, 9.60, 150.00),
        ("Letterpress birthday card", 6.00, 1.50, 60.00),
    ]
    for r, (item, price, var, fixed) in enumerate(rows, start=2):
        ws.append([item, price, var, fixed, f"=B{r}-C{r}",
                   f'=IF(E{r}<=0,"No break-even",ROUNDUP(D{r}/E{r},0))',
                   f'=IF(ISNUMBER(F{r}),F{r}*B{r},"")'])
        for col in "BCDEG":
            ws[f"{col}{r}"].number_format = USD
        for col in "EFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Break-Even Calculator for a Booth or Online Shop",
        "",
        "1. One row per product. Type Price, Variable cost per unit (materials, packaging, fees for one sale), and Fixed costs (booth fee, table rental, a one-time setup cost).",
        "2. Profit per unit = Price minus Variable cost per unit.",
        "3. Break-even units = ROUNDUP(Fixed costs / Profit per unit, 0). ROUNDUP, not ROUND, because you cannot sell part of a unit and be fully covered.",
        "4. Break-even revenue = Break-even units x Price.",
        "5. If Price is not higher than Variable cost per unit, Break-even units shows No break-even instead of a negative number.",
        "6. Copy the last row down to add products.",
        "",
        "Products, prices, and costs are fictional. This is arithmetic only; it does not include taxes or your own pay.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-break-even-calculator.xlsx")


def subscription_renewal_tracker():
    wb = Workbook()
    ws = wb.active
    ws.title = "Subscriptions"
    header(ws, ["Subscription", "Last renewed", "Months per cycle", "Cost per cycle", "Next renewal",
                "Days left", "Status", "Monthly cost", "Annual cost"],
           [28, 13, 11, 12, 13, 10, 14, 12, 12])
    rows = [
        ("Domain name", date(2026, 3, 15), 12, 18.00),
        ("Email marketing plan", date(2026, 9, 15), 1, 20.00),
        ("Online shop plan", date(2026, 9, 28), 1, 10.00),
        ("Bookkeeping software", date(2025, 11, 5), 12, 180.00),
        ("Photo storage (6 months)", date(2026, 8, 31), 6, 30.00),
    ]
    for r, (name, last, months, cost) in enumerate(rows, start=2):
        ws.append([name, last, months, cost,
                   f'=IF(B{r}="","",EDATE(B{r},C{r}))',
                   f'=IF(E{r}="","",E{r}-$K$1)',
                   f'=IF(F{r}="","",IF(F{r}<0,"Past due",IF(F{r}<=30,"Renews soon","OK")))',
                   f'=IF(D{r}="","",D{r}/C{r})',
                   f'=IF(D{r}="","",D{r}*12/C{r})'])
        for col in "BE":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "DHI":
            ws[f"{col}{r}"].number_format = USD
        for col in "EFGHI":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    last_row = 200
    ws["J1"], ws["K1"] = "As of date", date(2026, 10, 2)
    ws["J2"], ws["K2"] = "Total monthly", f"=SUM(H2:H{last_row})"
    ws["J3"], ws["K3"] = "Total annual", f"=SUM(I2:I{last_row})"
    ws["K1"].number_format = US_DATE
    ws["K2"].number_format = USD
    ws["K3"].number_format = USD
    for c in ("J1", "J2", "J3"):
        ws[c].font = Font(bold=True)
    ws.column_dimensions["J"].width = 16
    ws.column_dimensions["K"].width = 13
    ws.conditional_formatting.add(f"G2:G{last_row}", FormulaRule(formula=['OR($G2="Renews soon",$G2="Past due")'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Subscription Renewal Tracker",
        "",
        "1. One row per subscription. Type Last renewed (the date you were last charged), Months per cycle (1 = monthly, 12 = yearly), and Cost per cycle.",
        "2. Next renewal = EDATE(B2,C2), the same day of the month after Months per cycle. If that day does not exist (08/31 plus 6 months), EDATE uses the last day of the month: 02/28/2027.",
        "3. Days left = Next renewal minus the As of date in K1. Status shows Renews soon at 30 days or fewer, and Past due below 0.",
        "4. Monthly cost = Cost per cycle / Months per cycle. Annual cost = Cost per cycle x 12 / Months per cycle. Totals are in K2 and K3.",
        "5. The sample As of date is 10/02/2026 so the numbers stay put. Type =TODAY() in K1 to count from today.",
        "6. After a renewal, type the new charge date in Last renewed.",
        "",
        "Subscriptions and prices are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-subscription-renewal-tracker.xlsx")


def budget_vs_actual():
    wb = Workbook()
    ws = wb.active
    ws.title = "Budget vs actual"
    header(ws, ["Category", "Budget", "Actual", "Variance ($)", "Variance (%)", "Status"],
           [28, 12, 12, 13, 13, 14])
    rows = [
        ("Booth and market fees", 450.00, 450.00),
        ("Materials", 600.00, 683.40),
        ("Shipping", 220.00, 241.75),
        ("Packaging", 120.00, 131.20),
        ("Software subscriptions", 85.00, 85.00),
        ("Advertising", 150.00, 92.50),
        ("Payment processing fees", 90.00, 87.35),
        ("Training (not budgeted)", 0.00, 45.00),
    ]
    for r, (cat, bud, act) in enumerate(rows, start=2):
        ws.append([cat, bud, act, f"=C{r}-B{r}", f'=IF(B{r}=0,"",(C{r}-B{r})/B{r})',
                   f'=IF(C{r}>B{r},"Over budget","On or under")'])
        for col in "BCD":
            ws[f"{col}{r}"].number_format = USD
        ws[f"E{r}"].number_format = "0.0%"
        for col in "DEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    last = 200
    # Totals sit beside the table so rows can be added below without moving them.
    ws["H1"], ws["I1"] = "Total budget", f"=SUM(B2:B{last})"
    ws["H2"], ws["I2"] = "Total actual", f"=SUM(C2:C{last})"
    ws["H3"], ws["I3"] = "Total variance ($)", "=I2-I1"
    ws["H4"], ws["I4"] = "Total variance (%)", '=IF(I1=0,"",(I2-I1)/I1)'
    for c in ("I1", "I2", "I3"):
        ws[c].number_format = USD
    ws["I4"].number_format = "0.0%"
    for c in ("H1", "H2", "H3", "H4"):
        ws[c].font = Font(bold=True)
    ws.column_dimensions["H"].width = 20
    ws.column_dimensions["I"].width = 13
    ws.conditional_formatting.add(f"A2:F{last}", FormulaRule(formula=['AND($C2<>"",$C2>$B2)'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Budget vs Actual Variance",
        "",
        "1. One row per spending category. Type the Budget for the month and the Actual you spent.",
        "2. Variance ($) = Actual minus Budget. Positive means you spent more than planned.",
        "3. Variance (%) = (Actual - Budget) / Budget. It stays blank when Budget is 0, because dividing by zero is an error.",
        "4. Status says Over budget when Actual is more than Budget. Over-budget rows turn red by conditional formatting.",
        "5. Totals are in H1:I4, beside the table. Copy the last row down to add categories.",
        "",
        "Categories and amounts are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-budget-vs-actual.xlsx")


def farmers_market_sales_log():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Market date", "Item", "Qty sold", "Price", "Line total"], [13, 28, 10, 10, 12])
    d1, d2, d3 = date(2026, 9, 12), date(2026, 9, 19), date(2026, 9, 26)
    rows = [
        (d1, "Lavender soy candle, 8 oz", 9, 18.00),
        (d1, "Oatmeal goat milk soap bar", 14, 8.00),
        (d1, "Brass hoop earrings", 3, 24.00),
        (d1, "Letterpress birthday card", 12, 6.00),
        (d2, "Lavender soy candle, 8 oz", 11, 18.00),
        (d2, "Oatmeal goat milk soap bar", 10, 8.00),
        (d2, "Cedar soy candle, 8 oz", 6, 18.00),
        (d2, "Letterpress birthday card", 8, 6.00),
        (d2, "Brass hoop earrings", 2, 24.00),
        (d3, "Lavender soy candle, 8 oz", 7, 18.00),
        (d3, "Oatmeal goat milk soap bar", 16, 8.00),
        (d3, "Brass hoop earrings", 4, 24.00),
        (d3, "Letterpress birthday card", 10, 6.00),
    ]
    for r, (d, item, qty, price) in enumerate(rows, start=2):
        ws.append([d, item, qty, price, f"=C{r}*D{r}"])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = USD
        ws[f"E{r}"].number_format = USD
        ws[f"E{r}"].fill = FORMULA_FILL

    dt = wb.create_sheet("Daily totals")
    header(dt, ["Market date", "Revenue (SUMPRODUCT)", "Revenue (SUMIFS check)", "Items sold"],
           [13, 22, 22, 11])
    for r, d in enumerate((d1, d2, d3), start=2):
        dt.append([d,
                   f"=SUMPRODUCT((Sales!$A$2:$A$500=A{r})*Sales!$C$2:$C$500*Sales!$D$2:$D$500)",
                   f"=SUMIFS(Sales!$E$2:$E$500,Sales!$A$2:$A$500,A{r})",
                   f"=SUMIFS(Sales!$C$2:$C$500,Sales!$A$2:$A$500,A{r})"])
        dt[f"A{r}"].number_format = US_DATE
        dt[f"B{r}"].number_format = USD
        dt[f"C{r}"].number_format = USD
        for col in "BCD":
            dt[f"{col}{r}"].fill = FORMULA_FILL
    dt["F1"], dt["G1"] = "All days (SUMPRODUCT)", "=SUMPRODUCT(Sales!C2:C500,Sales!D2:D500)"
    dt["F2"], dt["G2"] = "All days (sum of Line total)", "=SUM(Sales!E2:E500)"
    dt["G1"].number_format = USD
    dt["G2"].number_format = USD
    dt.column_dimensions["F"].width = 28
    dt.column_dimensions["G"].width = 13
    for c in ("F1", "F2"):
        dt[c].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Farmers Market Sales Log",
        "",
        "1. Sales tab: one row per item sold at a market day. Type the Market date, Item, Qty sold, and Price. Line total = Qty x Price.",
        "2. Daily totals tab: type each market date in column A. Revenue (SUMPRODUCT) = SUMPRODUCT((Sales!$A$2:$A$500=A2)*Sales!$C$2:$C$500*Sales!$D$2:$D$500). It multiplies Qty by Price for that day's rows and adds them, without needing the Line total column.",
        "3. Revenue (SUMIFS check) adds the Line total column for that date. The two columns should match.",
        "4. G1 = SUMPRODUCT(Sales!C2:C500,Sales!D2:D500) is the grand total of Qty x Price across every day. G2 adds the Line total column; they should match too.",
        "5. Dates must be real dates, and Qty and Price must be numbers. Text in those cells makes the SUMPRODUCT formula in step 2 return #VALUE!.",
        "",
        "The vendor, items, and sales are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-farmers-market-sales-log.xlsx")


def clean_customer_list():
    wb = Workbook()
    ws = wb.active
    ws.title = "Clean names"
    header(ws, ["Messy name (paste here)", "Basic: PROPER(TRIM(CLEAN()))",
                "Full: also fixes non-breaking spaces and line breaks", "Length before", "Length after"],
           [30, 30, 40, 10, 10])
    nb = " "
    raw = [
        "  dana   ruiz ",
        "EMMA LEE",
        f"sam{nb}patel",
        "jamie\nchen",
        "  PRIYA   NAIR  ",
        "ana gomez",
        "ronald McDONALD",
        "sean o'neil",
        f"TOM  REYES{nb}",
        "   lee wu",
    ]
    for r, name in enumerate(raw, start=2):
        ws.append([name,
                   f"=PROPER(TRIM(CLEAN(A{r})))",
                   # UNICHAR(160), not CHAR(160): LibreOffice maps CHAR(160) through the
                   # system locale, so it misses non-breaking spaces under UTF-8. The
                   # _xlfn. prefix is how .xlsx files store UNICHAR (Excel 2013+, Sheets).
                   f'=PROPER(TRIM(CLEAN(SUBSTITUTE(SUBSTITUTE(A{r},_xlfn.UNICHAR(160)," "),CHAR(10)," "))))',
                   f"=LEN(A{r})", f"=LEN(C{r})"])
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        for col in "BCDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Clean Up a Customer List with TRIM, PROPER, and CLEAN",
        "",
        "1. Paste names in column A. Copy the gray formulas in B and C down next to them.",
        "2. TRIM removes extra spaces (leading, trailing, and doubled). CLEAN removes most hidden non-printing characters. PROPER capitalizes the first letter of each word.",
        "3. Column B is the basic formula =PROPER(TRIM(CLEAN(A2))). It does not remove non-breaking spaces (character 160, common in text copied from web pages), and CLEAN deletes a line break instead of replacing it with a space, so jamie + line break + chen becomes Jamiechen.",
        "4. Column C swaps both for normal spaces first: =PROPER(TRIM(CLEAN(SUBSTITUTE(SUBSTITUTE(A2,UNICHAR(160),\" \"),CHAR(10),\" \")))). UNICHAR(160) is used instead of CHAR(160) because LibreOffice can return a different character for CHAR(160) depending on the computer's locale.",
        "5. PROPER lowercases the rest of every word, so McDonald becomes Mcdonald. Fix names like that by hand.",
        "6. To keep the results, copy column C and Paste Special > Values only, then delete the formula columns.",
        "",
        "Names are fictional. The non-breaking space and line break behavior was checked in LibreOffice Calc only.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-clean-customer-list.xlsx")


def loan_payment_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Loans"
    header(ws, ["Loan", "Amount", "Yearly rate", "Years", "Monthly payment", "Total paid", "Total interest"],
           [28, 12, 11, 8, 15, 13, 14])
    rows = [
        ("Equipment loan (fictional)", 15000.00, 0.065, 5),
        ("Kiln purchase", 8000.00, 0.0725, 3),
        ("Booth trailer", 22000.00, 0.059, 7),
        ("Inventory line", 5000.00, 0.09, 2),
    ]
    for r, (name, amt, rate, yrs) in enumerate(rows, start=2):
        ws.append([name, amt, rate, yrs, f"=PMT(C{r}/12,D{r}*12,-B{r})", f"=E{r}*D{r}*12", f"=F{r}-B{r}"])
        ws[f"B{r}"].number_format = USD
        ws[f"C{r}"].number_format = "0.00%"
        for col in "EFG":
            ws[f"{col}{r}"].number_format = USD
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Loan Payment Calculator with PMT",
        "",
        "1. One row per loan. Type the Amount, the Yearly rate (6.5% or 0.065), and the Years.",
        "2. Monthly payment = PMT(rate/12, years*12, -amount). The rate is divided by 12 and the years are multiplied by 12 because the payments are monthly.",
        "3. The minus sign before the amount makes the payment come out positive. Without it PMT returns a negative number.",
        "4. Total paid = Monthly payment x number of payments. Total interest = Total paid minus Amount.",
        "5. Results can be a cent off a lender's statement because lenders round each month's payment.",
        "",
        "Loans and rates are fictional. This is arithmetic only, not financial advice.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-loan-payment-calculator.xlsx")


def percent_change():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Sales channel", "Last month", "This month", "Change ($)", "Change (%)"], [26, 13, 13, 13, 12])
    rows = [
        ("Whole shop", 2400.00, 2760.00),
        ("Etsy orders", 1800.00, 1692.00),
        ("Farmers market", 950.00, 1140.00),
        ("Workshops (new this month)", 0.00, 300.00),
        ("Wholesale", 640.00, 640.00),
    ]
    for r, (name, old, new) in enumerate(rows, start=2):
        ws.append([name, old, new, f"=C{r}-B{r}", f'=IF(B{r}=0,"",(C{r}-B{r})/B{r})'])
        for col in "BCD":
            ws[f"{col}{r}"].number_format = USD
        ws[f"E{r}"].number_format = "0.0%"
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Percent Change Between Two Months",
        "",
        "1. Type last month in column B and this month in column C.",
        "2. Change ($) = This month minus Last month.",
        "3. Change (%) = (This month - Last month) / Last month, formatted as Percent. Divide by the earlier month, not the later one.",
        "4. If Last month is 0 the percent is undefined, so the formula leaves the cell blank (see the Workshops row).",
        "5. A negative percent means the number went down.",
        "",
        "All figures are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-percent-change.xlsx")


def rank_top_customers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"
    header(ws, ["Customer", "Total sales", "Rank (ties share)", "Unique rank (ties broken)"], [26, 13, 14, 16])
    rows = [
        ("Harbor Coffee Co.", 4850.00), ("Maple Street Bakery", 3920.00), ("Oak & Ember Candles", 3920.00),
        ("Riverside Yoga", 2780.00), ("Juniper Florist", 2150.00), ("Bluebird Books", 2150.00),
        ("Cedar Hardware", 1640.00), ("Lakeview Dental Office", 3100.00), ("Pine Grove Garden", 980.00),
        ("North End Tailors", 1275.00),
    ]
    for r, (name, tot) in enumerate(rows, start=2):
        ws.append([name, tot, f"=RANK(B{r},$B$2:$B$11,0)", f"=C{r}+COUNTIF($B$2:B{r},B{r})-1"])
        ws[f"B{r}"].number_format = USD
        for col in "CD":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["F1"], ws["G1"], ws["H1"] = "Top 3", "Customer", "Total sales"
    for c in ("F1", "G1", "H1"):
        ws[c].fill, ws[c].font = HEADER_FILL, HEADER_FONT
    for i in range(1, 4):
        r = i + 1
        ws[f"F{r}"] = i
        ws[f"G{r}"] = f"=INDEX($A$2:$A$11,MATCH(F{r},$D$2:$D$11,0))"
        ws[f"H{r}"] = f"=INDEX($B$2:$B$11,MATCH(F{r},$D$2:$D$11,0))"
        ws[f"H{r}"].number_format = USD
        ws[f"G{r}"].fill = ws[f"H{r}"].fill = FORMULA_FILL
    ws.column_dimensions["F"].width = 8
    ws.column_dimensions["G"].width = 26
    ws.column_dimensions["H"].width = 13

    notes_sheet(wb, [
        "Tidy Tabs: Rank Customers by Sales",
        "",
        "1. Type customer names in A and total sales in B.",
        "2. Rank = RANK(B2,$B$2:$B$11,0). The 0 ranks the biggest total as 1. Use 1 instead of 0 to rank the smallest as 1.",
        "3. Tied totals get the same rank, and the next rank is skipped (two customers at rank 2 means no rank 3).",
        "4. Unique rank = Rank + COUNTIF($B$2:B2,B2) - 1. The first customer with a total keeps the rank, the next one with the same total gets rank + 1.",
        "5. The Top 3 block uses INDEX/MATCH on the unique rank, so it never hits a missing rank.",
        "6. Lock the range with $ signs ($B$2:$B$11) so it does not slide when you fill down. Extend it when you add customers.",
        "",
        "Customers and totals are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-rank-top-customers.xlsx")


def invoice_number_generator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Client", "Invoice date", "Invoice number (TEXT)", "Invoice number (YEAR)", "Amount"],
           [26, 13, 22, 22, 12])
    rows = [
        ("Harbor Coffee Co.", date(2026, 9, 28), 450.00),
        ("Maple Street Bakery", date(2026, 9, 30), 1200.00),
        ("Riverside Yoga", date(2026, 10, 2), 315.00),
        ("Juniper Florist", date(2026, 10, 5), 780.00),
        ("Bluebird Books", date(2026, 10, 9), 95.00),
        ("Cedar Hardware", date(2026, 10, 12), 640.00),
    ]
    for r, (client, d, amt) in enumerate(rows, start=2):
        ws.append([client, d,
                   f'="INV-"&TEXT(B{r},"yyyy")&"-"&TEXT(ROWS($A$2:A{r}),"0000")',
                   f'="INV-"&YEAR(B{r})&"-"&TEXT(ROWS($A$2:A{r}),"0000")', amt])
        ws[f"B{r}"].number_format = US_DATE
        ws[f"E{r}"].number_format = USD
        for col in "CD":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Automatic Invoice Numbers",
        "",
        "1. Type the client and invoice date. Copy the gray formulas down for each new row.",
        "2. Column C: =\"INV-\"&TEXT(B2,\"yyyy\")&\"-\"&TEXT(ROWS($A$2:A2),\"0000\"). TEXT(...,\"0000\") pads the count to four digits; ROWS counts how many rows are in the range, so each new row gets the next number.",
        "3. Column D does the year with YEAR(B2) instead of a TEXT format code. Use it if a date format code such as yyyy does not work in your language or locale.",
        "4. The number follows the row position, not a saved counter. If you sort or delete rows, the numbers shift. Once an invoice is sent, copy the numbers and Paste Special > Values only so they stay fixed.",
        "5. The year comes from the invoice date, so the count does not restart on its own each January. Start a new sheet or tab each year.",
        "",
        "Clients and amounts are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-invoice-number-generator.xlsx")


def combine_address_columns():
    wb = Workbook()
    ws = wb.active
    ws.title = "Addresses"
    header(ws, ["Street", "Unit", "City", "State", "ZIP", "TEXTJOIN (skips blanks)", "With ZIP", "Plain & join (stray commas)"],
           [22, 9, 14, 7, 8, 38, 44, 38])
    rows = [
        ("118 Elm Street", "Apt 4B", "Boston", "MA", "02108"),
        ("2450 Harbor Road", "", "Portland", "ME", "04101"),
        ("77 Mill Lane", "Suite 210", "Hoboken", "NJ", "07030"),
        ("910 Cedar Avenue", "", "Providence", "RI", "02903"),
        ("36 Orchard Way", "Unit 2", "Albany", "NY", "12207"),
    ]
    for r, row in enumerate(rows, start=2):
        ws.append(list(row) + [f'=_xlfn.TEXTJOIN(", ",TRUE,A{r}:D{r})',
                               f'=_xlfn.TEXTJOIN(", ",TRUE,A{r}:D{r})&" "&E{r}',
                               f'=A{r}&", "&B{r}&", "&C{r}&", "&D{r}'])
        ws[f"E{r}"].number_format = "@"
        for col in "FGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Combine Address Columns with TEXTJOIN",
        "",
        "1. Put street, unit, city, and state in A to D and the ZIP in E (formatted as Text so 02108 keeps its zero).",
        "2. Column F: =TEXTJOIN(\", \",TRUE,A2:D2). The first argument is the separator, TRUE skips empty cells, and the last is the range to join.",
        "3. Column G adds the ZIP after a space, so the state and ZIP are not split by a comma.",
        "4. Column H shows the older way, joining with &. When Unit is empty it leaves a double comma, which is what TEXTJOIN avoids.",
        "5. TEXTJOIN is in Excel 2019 and Microsoft 365, and in Google Sheets. It is not in Excel 2016 or earlier.",
        "",
        "Addresses are fictional. Checked in LibreOffice Calc only.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-combine-address-columns.xlsx")


def sale_price_discount():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sale prices"
    header(ws, ["Item", "Original price", "Percent off", "Sale price", "You save"],
           [30, 14, 12, 12, 12])
    rows = [
        ("Lavender soy candle, 8 oz", 18.00, 0.10),
        ("Cedar soy candle, 8 oz", 18.00, 0.15),
        ("Oatmeal goat milk soap bar", 8.00, 0.25),
        ("Charcoal soap bar", 8.50, 0.10),
        ("Brass hoop earrings", 24.00, 0.15),
        ("Letterpress birthday card", 6.00, 0.25),
        ("Beaded necklace", 48.00, 0.25),
        ("Candle and soap gift set", 32.00, 0.10),
    ]
    for r, (item, price, pct) in enumerate(rows, start=2):
        ws.append([item, price, pct, f"=B{r}*(1-C{r})", f"=B{r}-D{r}"])
        for col in "BDE":
            ws[f"{col}{r}"].number_format = USD
        ws[f"C{r}"].number_format = "0%"
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["G1"], ws["H1"] = "Total at original prices", "=SUM(B2:B500)"
    ws["G2"], ws["H2"] = "Total at sale prices", "=SUM(D2:D500)"
    ws["G3"], ws["H3"] = "Total saved", "=SUM(E2:E500)"
    for r in (1, 2, 3):
        ws[f"G{r}"].font = Font(bold=True)
        ws[f"H{r}"].number_format = USD
    ws.column_dimensions["G"].width = 24
    ws.column_dimensions["H"].width = 13

    rv = wb.create_sheet("Reverse and stacked")
    header(rv, ["Item", "Where you saw it", "Percent off", "Sale price", "Original price"],
           [28, 18, 12, 12, 14])
    for r, (item, where, pct, sale) in enumerate([
        ("Beaded necklace", "Market tag", 0.25, 36.00),
        ("Brass hoop earrings", "Etsy listing", 0.15, 20.40),
        ("Charcoal soap bar", "Receipt", 0.10, 7.65),
    ], start=2):
        rv.append([item, where, pct, sale, f"=D{r}/(1-C{r})"])
        rv[f"C{r}"].number_format = "0%"
        for col in "DE":
            rv[f"{col}{r}"].number_format = USD
        rv[f"E{r}"].fill = FORMULA_FILL
    rv["A7"] = "Two discounts in a row"
    rv["A7"].font = Font(bold=True, size=12)
    labels = [
        ("Original price", 100.00),
        ("First discount", 0.20),
        ("Second discount", 0.10),
        ("Price after both, one at a time", "=B8*(1-B9)*(1-B10)"),
        ("Wrong: add the percents (30% off)", "=B8*(1-(B9+B10))"),
        ("Difference", "=B11-B12"),
        ("Real total discount", "=1-B11/B8"),
    ]
    for r, (label, val) in enumerate(labels, start=8):
        rv[f"A{r}"], rv[f"B{r}"] = label, val
    for r in (8, 11, 12, 13):
        rv[f"B{r}"].number_format = USD
    for r in (9, 10, 14):
        rv[f"B{r}"].number_format = "0%"
    for r in (11, 12, 13, 14):
        rv[f"B{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Sale Price and Discount Calculator",
        "",
        "1. Sale prices tab: type the Original price and Percent off. Sale price = B2*(1-C2). You save = B2-D2.",
        "2. Format Percent off as a percentage and type 25% (or 0.25), not 25.",
        "3. Reverse and stacked tab, top: you know the sale price and the percent off and want the original price. Original price = D2/(1-C2).",
        "4. Reverse and stacked tab, bottom: 20% off and then 10% off is not 30% off. $100.00 becomes $80.00, then $72.00, a real discount of 28%.",
        "5. Totals for the item list are in H1:H3.",
        "",
        "Items and prices are fictional. This is arithmetic only; it does not include sales tax or shipping.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-sale-price-discount-calculator.xlsx")


def shipping_weight_tier_lookup():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order #", "Weight (oz)", "Rate (VLOOKUP)", "Rate (INDEX/MATCH)", "VLOOKUP without IFERROR",
                "", "From (oz)", "Rate"],
           [10, 12, 15, 18, 22, 4, 11, 10])
    # Lower edge of each band, smallest to largest. An exact hit on an edge
    # belongs to that band, so 4 oz is in the 4 oz band.
    table = [(1, 4.50), (4, 5.75), (8, 7.25), (16, 9.80), (32, 13.40), (64, 18.90)]
    weights = [2.5, 3.99, 4, 4.01, 7.9, 8, 12.5, 16, 20, 40, 64, 0.5]
    for r, w in enumerate(weights, start=2):
        ws.append([str(2100 + r - 1), w,
                   f'=IFERROR(VLOOKUP(B{r},$G$2:$H$7,2,TRUE),"Below minimum")',
                   f'=IFERROR(INDEX($H$2:$H$7,MATCH(B{r},$G$2:$G$7,1)),"Below minimum")',
                   f"=VLOOKUP(B{r},$G$2:$H$7,2,TRUE)"])
        for col in "CDE":
            ws[f"{col}{r}"].number_format = USD
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"B{r}"].number_format = "0.00"
    for r, (lo, rate) in enumerate(table, start=2):
        ws[f"G{r}"], ws[f"H{r}"] = lo, rate
        ws[f"H{r}"].number_format = USD
    for c in ("G1", "H1"):
        ws[c].fill, ws[c].font = HEADER_FILL, HEADER_FONT

    un = wb.create_sheet("Unsorted table demo")
    header(un, ["Weight (oz)", "VLOOKUP result", "Correct rate (sorted table)", "", "From (oz)", "Rate"],
           [12, 16, 24, 4, 11, 10])
    unsorted = [(1, 4.50), (8, 7.25), (4, 5.75), (16, 9.80), (32, 13.40), (64, 18.90)]
    for r, w in enumerate([5, 10], start=2):
        un.append([w, f"=VLOOKUP(A{r},$E$2:$F$7,2,TRUE)",
                   f"=VLOOKUP(A{r},Orders!$G$2:$H$7,2,TRUE)"])
        for col in "BC":
            un[f"{col}{r}"].number_format = USD
        un[f"B{r}"].fill = ALERT_FILL
    for r, (lo, rate) in enumerate(unsorted, start=2):
        un[f"E{r}"], un[f"F{r}"] = lo, rate
        un[f"F{r}"].number_format = USD

    notes_sheet(wb, [
        "Tidy Tabs: Shipping Cost by Weight Tier (approximate match lookup)",
        "",
        "1. Rate table in G:H, Orders tab. From (oz) is the smallest weight in each band. A package at or above that weight and below the next row uses that rate.",
        "2. Rate (VLOOKUP) = IFERROR(VLOOKUP(B2,$G$2:$H$7,2,TRUE),\"Below minimum\"). TRUE (or leaving it out) means approximate match: the largest From value that is less than or equal to the weight.",
        "3. Rate (INDEX/MATCH) does the same with MATCH(...,1). Both columns agree.",
        "4. The From column must be sorted smallest to largest. Approximate match skips through the list instead of checking every row, so an unsorted table can return a wrong rate with no error. See the Unsorted table demo tab: in LibreOffice 5 oz happens to come out right ($5.75) but 10 oz returns $5.75 instead of $7.25. Other apps can pick differently on an unsorted table.",
        "5. A weight under the first band (0.50 oz here) has no band, so VLOOKUP returns #N/A. Column E shows it; IFERROR in column C turns it into text. IFERROR also hides other errors, so test the table first.",
        "6. Exactly 4 oz is in the 4 oz band ($5.75). 3.99 oz is in the 1 oz band ($4.50). 4.01 oz is in the 4 oz band.",
        "",
        "Rates and weight bands are made up for this sample, not a carrier's rates. Use your own.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-shipping-weight-tier-lookup.xlsx")


def reorder_point_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Reorder"
    header(ws, ["Supply", "Units sold (last 30 days)", "Daily sales", "Lead time (days)", "Safety stock",
                "Reorder point", "On hand", "Flag"],
           [28, 14, 11, 12, 11, 12, 10, 12])
    rows = [
        ("Glass jars, 8 oz", 90, 7, 6, 40),
        ("Soy wax, 1 lb bags", 60, 10, 8, 28),
        ("Candle wicks, pack of 50", 33, 14, 4, 21),
        ("Cotton ribbon spools", 75, 5, 10, 18),
        ("Kraft mailer boxes", 120, 12, 15, 80),
        ("Shipping label rolls", 45, 21, 5, 30),
    ]
    for r, (item, sold, lead, safety, onhand) in enumerate(rows, start=2):
        ws.append([item, sold, f"=B{r}/$K$1", lead, safety,
                   f"=ROUNDUP(C{r}*D{r}+E{r},0)", onhand,
                   f'=IF(G{r}<=F{r},"REORDER","OK")'])
        ws[f"C{r}"].number_format = "0.00"
        for col in "CFH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    last = 200
    ws["J1"], ws["K1"] = "Days in sales period", 30
    ws["J1"].font = Font(bold=True)
    ws.column_dimensions["J"].width = 22
    ws.conditional_formatting.add(f"H2:H{last}", FormulaRule(formula=['$H2="REORDER"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Reorder Point Calculator",
        "",
        "1. One row per supply. Type Units sold for the last 30 days, Lead time (days) from order to delivery, Safety stock (spare units you want to keep), and On hand.",
        "2. Daily sales = Units sold / days in the period. The period length is in K1 (30).",
        "3. Reorder point = ROUNDUP(Daily sales x Lead time + Safety stock, 0). ROUNDUP, because a fraction of a unit still needs a whole unit.",
        "4. Flag = IF(On hand <= Reorder point, \"REORDER\", \"OK\"). At exactly the reorder point the flag says REORDER.",
        "5. Example: 90 units in 30 days is 3 a day. 3 x 7 lead days + 6 safety stock = 27.",
        "6. Copy the last row down to add supplies.",
        "",
        "Supplies and numbers are fictional. Safety stock is your own judgment; this file only does the arithmetic.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-reorder-point-calculator.xlsx")


def freelance_hourly_rate():
    wb = Workbook()
    ws = wb.active
    ws.title = "Hourly rate"
    header(ws, ["", "Your numbers", "What if 1: higher goal", "What if 2: more time off", "What if 3: more billable"],
           [36, 15, 18, 18, 18])
    inputs = [
        ("Yearly income goal", 60000, 75000, 60000, 60000, USD),
        ("Yearly overhead (software, supplies)", 6000, 6000, 6000, 6000, USD),
        ("Weeks off per year", 4, 4, 6, 4, "0"),
        ("Hours per week you work", 30, 30, 30, 30, "0"),
        ("Share of those hours you bill", 0.65, 0.65, 0.65, 0.75, "0%"),
    ]
    for r, (label, *vals, fmt) in enumerate(inputs, start=2):
        ws.append([label] + vals)
        for col in "BCDE":
            ws[f"{col}{r}"].number_format = fmt
    calc = [
        ("Weeks worked", "=52-{c}4", "0"),
        ("Hours worked per year", "={c}7*{c}5", "#,##0"),
        ("Billable hours per year", "={c}8*{c}6", "#,##0"),
        ("Money you need to bring in", "={c}2+{c}3", USD),
        ("Hourly rate", "={c}10/{c}9", USD),
        ("Rate rounded up to a whole dollar", "=ROUNDUP({c}11,0)", USD),
    ]
    for r, (label, f, fmt) in enumerate(calc, start=7):
        ws[f"A{r}"] = label
        for col in "BCDE":
            ws[f"{col}{r}"] = f.format(c=col)
            ws[f"{col}{r}"].number_format = fmt
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["A11"].font = Font(bold=True)
    for col in "BCDE":
        ws[f"{col}11"].font = Font(bold=True)
    ws.freeze_panes = "B2"

    notes_sheet(wb, [
        "Tidy Tabs: Freelance Hourly Rate Calculator",
        "",
        "1. Type your numbers in column B, rows 2 to 6. Columns C to E are what-if copies: change any input there and the rate below it updates.",
        "2. Weeks worked = 52 - weeks off. Hours worked = weeks worked x hours per week. Billable hours = hours worked x billable share.",
        "3. Money you need = income goal + overhead. Hourly rate = money you need / billable hours.",
        "4. Example: (52-4) x 30 x 65% = 936 billable hours. ($60,000.00 + $6,000.00) / 936 = $70.51 an hour.",
        "5. Billable share is the part of your work week you can actually charge for. Time spent finding clients, invoicing and emailing is not billable.",
        "",
        "Illustration only. This does not cover taxes, self-employment costs, or benefits; ask an accountant about those. All numbers are made up.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-freelance-hourly-rate-calculator.xlsx")


def convert_text_to_numbers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Pasted amounts"
    header(ws, ["Pasted text", "Is it a number?", "Basic clean", "Clean with (negatives)", "Is it a number now?"],
           [16, 14, 14, 22, 18])
    texts = ["$1,250.00", " 18.00", "(45.50)", "$89.95", "$2,400.00 ", "12.50",
             "$7.25", "$1,075.40", "$310.00", "(22.00)"]
    for r, t in enumerate(texts, start=2):
        ws.append([t, f"=ISNUMBER(A{r})",
                   f'=VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A{r}),"$",""),",",""))',
                   f'=IF(LEFT(TRIM(A{r}),1)="(",-1,1)*VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A{r}),"$",""),",",""),"(",""),")",""))',
                   f"=ISNUMBER(D{r})"])
        ws[f"A{r}"].number_format = "@"
        ws[f"A{r}"].alignment = Alignment(horizontal="left")
        for col in "CD":
            ws[f"{col}{r}"].number_format = USD
        for col in "BCDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    labels = [("SUM of the pasted text", "=SUM(A2:A500)"),
              ("SUM of Basic clean", "=SUM(C2:C500)"),
              ("SUM of Clean with (negatives)", "=SUM(D2:D500)")]
    for r, (label, f) in enumerate(labels, start=1):
        ws[f"G{r}"], ws[f"H{r}"] = label, f
        ws[f"G{r}"].font = Font(bold=True)
        ws[f"H{r}"].number_format = USD
    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["H"].width = 14

    notes_sheet(wb, [
        "Tidy Tabs: Convert Text to Numbers (pasted dollar amounts)",
        "",
        "1. Column A holds amounts as text, the way they arrive when you paste from a bank or Etsy export. SUM ignores text, so H1 is $0.00.",
        "2. Column B: =ISNUMBER(A2) is FALSE for text and TRUE for a real number.",
        "3. Column C: =VALUE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),\"$\",\"\"),\",\",\"\")). TRIM removes stray spaces, SUBSTITUTE removes $ and commas, VALUE turns the text into a number.",
        "4. Column D also handles amounts in parentheses, such as (45.50), as negatives: it removes the ( and ) and multiplies by -1. LibreOffice reads (45.50) as -45.50 even in column C, but do not count on that in other apps; column D does not depend on it.",
        "5. Column E checks that the cleaned values are real numbers. Copy C or D, then Paste Special > Values to keep the numbers.",
        "",
        "Amounts are fictional. Checked in LibreOffice Calc with a US locale only; other regional settings can read $, commas and periods differently.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-convert-text-to-numbers.xlsx")


def sales_tax_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales tax"
    header(ws, ["Item", "Price before tax", "Sales tax", "Total"], [28, 16, 12, 12])
    rows = [
        ("Ceramic mug", 18.00),
        ("Sticker pack", 4.50),
        ("Canvas tote bag", 24.00),
        ("Soy candle, 8 oz", 16.00),
        ("Art print, 8 x 10", 12.99),
        ("Beaded earrings", 32.50),
    ]
    for r, (item, price) in enumerate(rows, start=2):
        ws.append([item, price, f"=ROUND(B{r}*$F$1,2)", f"=B{r}+C{r}"])
        for col in "BCD":
            ws[f"{col}{r}"].number_format = USD
        for col in "CD":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["A8"] = "Total"
    ws["A8"].font = Font(bold=True)
    for col in "BCD":
        ws[f"{col}8"] = f"=SUM({col}2:{col}7)"
        ws[f"{col}8"].number_format = USD
        ws[f"{col}8"].font = Font(bold=True)
    ws["E1"], ws["F1"] = "Sales tax rate", 0.0625
    ws["E1"].font = Font(bold=True)
    ws["F1"].number_format = "0.00%"
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 10

    inc = wb.create_sheet("Tax-inclusive totals")
    header(inc, ["Item", "Total paid (tax included)", "Tax inside the total", "Price before tax"], [28, 16, 16, 16])
    paid = [
        ("Ceramic mug", 19.13),
        ("Sticker pack", 4.78),
        ("Canvas tote bag", 25.50),
        ("Soy candle, 8 oz", 17.00),
        ("Art print, 8 x 10", 13.80),
        ("Beaded earrings", 34.53),
        ("Gift set, one price at a booth", 50.00),
    ]
    for r, (item, total) in enumerate(paid, start=2):
        inc.append([item, total, f"=ROUND(B{r}-B{r}/(1+$F$1),2)", f"=B{r}-C{r}"])
        for col in "BCD":
            inc[f"{col}{r}"].number_format = USD
        for col in "CD":
            inc[f"{col}{r}"].fill = FORMULA_FILL
    inc["E1"], inc["F1"] = "Sales tax rate", "='Sales tax'!F1"
    inc["E1"].font = Font(bold=True)
    inc["F1"].number_format = "0.00%"
    inc.column_dimensions["E"].width = 16
    inc.column_dimensions["F"].width = 10

    notes_sheet(wb, [
        "Tidy Tabs: Sales Tax Calculator",
        "",
        "1. Type your rate in F1 on the Sales tax tab (6.25% is a made-up example). Rates vary by state, county and city, so check your state's revenue department site for the real rate.",
        "2. Sales tax = ROUND(Price x rate, 2). ROUND to cents because you cannot collect a fraction of a cent.",
        "3. Total = Price + Sales tax. Row 8 adds up each column.",
        "4. Tax-inclusive totals tab: type what the customer paid in column B. Tax inside = ROUND(Total - Total / (1 + rate), 2). Price before tax = Total - Tax.",
        "5. The rate on the second tab follows F1 on the first tab.",
        "6. Example: $18.00 x 6.25% = $1.125, rounded to $1.13, so the total is $19.13.",
        "",
        "Items, prices and the rate are fictional. This sheet only does arithmetic and is not tax advice.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-sales-tax-calculator.xlsx")




def sum_expenses_by_category():
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    header(ws, ["Date", "Vendor", "Category", "Amount", "", "Category", "Total spent", "% of total"],
           [12, 24, 13, 12, 3, 16, 14, 11])
    log = [
        (date(2026, 9, 1), "Greenline Office Supply", "Supplies", 84.50),
        (date(2026, 9, 2), "Lakeview Web Hosting", "Software", 29.00),
        (date(2026, 9, 3), "Pioneer Postal Center", "Shipping", 46.25),
        (date(2026, 9, 5), "Metro Transit Pass", "Travel", 90.00),
        (date(2026, 9, 6), "Harbor Coffee", "Meals", 18.40),
        (date(2026, 9, 8), "Northeast Electric", "Utilities", 112.30),
        (date(2026, 9, 9), "Social Ads Co", "Advertising", 150.00),
        (date(2026, 9, 10), "Greenline Office Supply", "Supplies", 37.80),
        (date(2026, 9, 12), "Pioneer Postal Center", "Shipping", 62.10),
        (date(2026, 9, 13), "Cloud Backup Plus", "Software", 12.99),
        (date(2026, 9, 15), "Oak Street Deli", "Meals", 24.75),
        (date(2026, 9, 16), "Harbor Parking Garage", "Travel", 22.00),
        (date(2026, 9, 17), "Social Ads Co", "Advertising", 75.00),
        (date(2026, 9, 19), "Craft Paper Wholesale", "Supplies", 128.60),
        (date(2026, 9, 20), "Northeast Gas", "Utilities", 58.45),
        (date(2026, 9, 22), "Pioneer Postal Center", "Shipping", 33.90),
        (date(2026, 9, 23), "Design Template Shop", "Software", 49.00),
        (date(2026, 9, 24), "Amtrak Regional", "Travel", 134.00),
        (date(2026, 9, 25), "Maple Cafe", "Meals", 31.20),
        (date(2026, 9, 26), "Craft Paper Wholesale", "Supplies", 66.40),
        (date(2026, 9, 27), "Local Print Shop", "Advertising", 95.50),
        (date(2026, 9, 29), "Lakeview Internet", "Utilities", 69.99),
    ]
    for r, (d, vendor, cat, amt) in enumerate(log, start=2):
        ws.append([d, vendor, cat, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = USD
    cats = ["Supplies", "Software", "Shipping", "Travel", "Meals", "Utilities", "Advertising"]
    for r, cat in enumerate(cats, start=2):
        ws[f"F{r}"] = cat
        ws[f"G{r}"] = f"=SUMIF($C$2:$C$40,F{r},$D$2:$D$40)"
        ws[f"H{r}"] = f"=G{r}/$G$9"
        ws[f"G{r}"].number_format = USD
        ws[f"H{r}"].number_format = "0.0%"
        ws[f"G{r}"].fill = ws[f"H{r}"].fill = FORMULA_FILL
    ws["F9"], ws["G9"], ws["H9"] = "Total", "=SUM(G2:G8)", "=SUM(H2:H8)"
    ws["G9"].number_format, ws["H9"].number_format = USD, "0.0%"
    ws["F11"], ws["G11"] = "Log total", "=SUM($D$2:$D$40)"
    ws["F12"], ws["G12"] = "Difference", "=ROUND(G11-G9,2)"
    ws["F13"], ws["G13"] = "Check", '=IF(G12=0,"OK","CHECK CATEGORIES")'
    for c in ("F9", "G9", "H9", "F11", "F12", "F13"):
        ws[c].font = Font(bold=True)
    ws["G11"].number_format = ws["G12"].number_format = USD
    ws["G11"].fill = ws["G12"].fill = ws["G13"].fill = FORMULA_FILL
    ws.conditional_formatting.add("G13", FormulaRule(formula=['$G$13<>"OK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Sum Expenses by Category",
        "",
        "1. Log one expense per row in A to D (Date, Vendor, Category, Amount). The formulas read rows 2 to 40, so you can add up to 39 rows without editing anything.",
        "2. Type each category exactly the same way every time. The category list is in F2:F8.",
        "3. Total spent = SUMIF($C$2:$C$40,F2,$D$2:$D$40). It adds column D wherever column C matches the category in column F.",
        "4. % of total = G2/$G$9, where G9 is the sum of the category totals.",
        "5. Check: G11 = SUM($D$2:$D$40), every amount in the log. G12 is G11 minus G9. If it is not $0.00, G13 says CHECK CATEGORIES and some expense has a category that is not in the list.",
        "6. Common cause: a typo or a trailing space, such as \"Supplies \" with a space after it. Retype the category, or clean the column with TRIM.",
        "7. To add a category, insert a row between rows 2 and 8 (so the Total row still covers it) and copy the formulas from the row above.",
        "",
        "Vendors and amounts are fictional. This file only adds up what you type; it is not tax or accounting advice.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-sum-expenses-by-category.xlsx")


def round_prices_nearest_nickel_99():
    wb = Workbook()
    ws = wb.active
    ws.title = "Prices"
    header(ws, ["Item", "Price after markup math", "Nearest cent", "Nearest $0.05", "Up to next $0.05",
                "Ends in .99"],
           [30, 16, 13, 14, 15, 13])
    rows = [
        ("Beeswax candle, 4 oz", 13.4167),
        ("Cotton tote bag", 8.2333),
        ("Hand-poured soap set", 24.7083),
        ("Ceramic mug", 17.9625),
        ("Greeting card, single", 6.13),
        ("Linen napkins, set of 4", 31.4467),
        ("Sticker sheet", 12.9833),
        ("Embroidered patch", 22.0),
    ]
    for r, (item, price) in enumerate(rows, start=2):
        ws.append([item, price, f"=ROUND(B{r},2)", f"=MROUND(B{r},0.05)", f"=CEILING(B{r},0.05)",
                   f"=ROUNDUP(B{r},0)-0.01"])
        ws[f"B{r}"].number_format = '"$"#,##0.0000'
        for col in "CDEF":
            ws[f"{col}{r}"].number_format = USD
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Round Prices to the Nearest $0.05 or .99",
        "",
        "1. Type your unrounded prices in column B (the result of your markup or discount math).",
        "2. Nearest cent: =ROUND(B2,2). Nearest nickel: =MROUND(B2,0.05). Always up to the next nickel: =CEILING(B2,0.05).",
        "3. Ending in .99: =ROUNDUP(B2,0)-0.01. It rounds up to the next whole dollar, then takes off a cent.",
        "4. Watch out: a price that is already a whole dollar, such as $22.00, becomes $21.99 with the .99 formula.",
        "5. Example: $13.4167 becomes $13.42 (cent), $13.40 (nickel), $13.45 (up) and $13.99 (.99).",
        "6. Copy the last row down to add items.",
        "",
        "Items and prices are fictional. Checked in LibreOffice Calc only.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-round-prices-nearest-nickel-99.xlsx")


def remove_non_breaking_spaces():
    wb = Workbook()
    ws = wb.active
    ws.title = "Pasted names"
    header(ws, ["Pasted name", "LEN before", "Code at position 5", "TRIM only", "Cleaned",
                "LEN after", "COUNTIF before", "COUNTIF after", "Plan price before", "Plan price after",
                "", "Customer list", "Plan price"],
           [18, 10, 12, 14, 14, 10, 12, 12, 14, 14, 3, 16, 12])
    nb = chr(160)  # built in with Python so the cells hold a real non-breaking space
    raw = [
        f"Dana{nb}Ruiz",
        f"Emma{nb}Lee",
        f"Sara{nb}Kim{nb}",
        "Mark Diaz\n",
        f"Lena{nb}Ortiz\t",
        f"Josh{nb}{nb}Park",
        f"Anna{nb}Cho",
        "Mike Stone",
    ]
    book = [("Dana Ruiz", 120), ("Emma Lee", 45), ("Sara Kim", 89.5), ("Mark Diaz", 210),
            ("Lena Ortiz", 65), ("Josh Park", 150), ("Anna Cho", 38), ("Mike Stone", 72.25)]
    for r, name in enumerate(raw, start=2):
        ws.append([name, f"=LEN(A{r})", f"=_xlfn.UNICODE(MID(A{r},5,1))", f"=TRIM(A{r})",
                   # UNICHAR/UNICODE, not CHAR/CODE: LibreOffice maps CHAR(160) and CODE through
                   # the system locale (seen: 160 in one profile, 194 in a fresh one).
                   f'=TRIM(CLEAN(SUBSTITUTE(A{r},_xlfn.UNICHAR(160)," ")))', f"=LEN(E{r})",
                   f"=COUNTIF($L$2:$L$9,A{r})", f"=COUNTIF($L$2:$L$9,E{r})",
                   f'=IFERROR(VLOOKUP(A{r},$L$2:$M$9,2,FALSE),"not found")',
                   f'=IFERROR(VLOOKUP(E{r},$L$2:$M$9,2,FALSE),"not found")'])
        ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        for col in "BCDEFGHIJ":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        for col in "IJ":
            ws[f"{col}{r}"].number_format = USD
            ws[f"{col}{r}"].alignment = Alignment(horizontal="right")
        ws[f"L{r}"], ws[f"M{r}"] = book[r - 2]
        ws[f"M{r}"].number_format = USD
    ws["L11"], ws["M11"] = "Found before", "=SUM(G2:G9)"
    ws["L12"], ws["M12"] = "Found after", "=SUM(H2:H9)"
    for c in ("L11", "L12"):
        ws[c].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Remove Non-Breaking Spaces (CHAR 160) in Excel and Google Sheets",
        "",
        "1. Column A holds fictional names pasted the way web pages and PDFs deliver them: most have a non-breaking space (character 160) where a normal space should be, and some have a trailing line break or tab.",
        "2. Column B: =LEN(A2) counts every character, including the invisible ones. Dana Ruiz shows 9 instead of 8.",
        "3. Column C: =UNICODE(MID(A2,5,1)) shows the character code at position 5 (=CODE(MID(A2,5,1)) is the older version, but in LibreOffice it returned 160 in one setup and 194 in another, depending on the locale). A normal space is 32, a non-breaking space is 160. This works here because every first name has 4 letters; change the 5 to point at the character you suspect.",
        "4. Column D: =TRIM(A2) only removes the regular space character (code 32), so the non-breaking spaces stay. Compare its length to column B.",
        "5. Column E: =TRIM(CLEAN(SUBSTITUTE(A2,UNICHAR(160),\" \"))). UNICHAR(160) is used instead of CHAR(160) because LibreOffice returned a different character for CHAR(160) in one setup. SUBSTITUTE swaps character 160 for a normal space, CLEAN removes line breaks and tabs, TRIM removes extra spaces.",
        "6. CLEAN deletes a line break or tab instead of replacing it with a space. If one sits between two words, swap it first: SUBSTITUTE(A2,CHAR(10),\" \").",
        "7. Columns G to J look each name up in the customer list in L:M. Before cleaning, only Mike Stone is found. After cleaning, all 8 are found.",
        "8. To keep the results, copy column E and Paste Special > Values only, then delete the gray formula columns.",
        "",
        "Names and prices are fictional. Checked in LibreOffice Calc 24.2.7 (Linux) only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-remove-non-breaking-spaces.xlsx")


def year_to_date_sales_total():
    rows = [
        (date(2026, 1, 12), "Harbor Coffee", 1250.00),
        (date(2026, 1, 28), "Maple Street Bakery", 480.00),
        (date(2026, 2, 10), "Oak & Ember Candles", 920.50),
        (date(2026, 2, 24), "Harbor Coffee", 1100.00),
        (date(2026, 4, 7), "Maple Street Bakery", 640.00),
        (date(2026, 5, 19), "Harbor Coffee", 1380.75),
        (date(2026, 6, 23), "Oak & Ember Candles", 705.00),
        (date(2026, 7, 8), "Lakeview Florist", 450.00),
        (date(2026, 7, 30), "Harbor Coffee", 1015.00),
        (date(2026, 8, 4), "Maple Street Bakery", 520.50),
        (date(2026, 8, 18), "Oak & Ember Candles", 860.00),
        (date(2026, 8, 31), "Lakeview Florist", 275.00),
        (date(2026, 9, 12), "Harbor Coffee", 1190.00),
        (date(2026, 10, 3), "Lakeview Florist", 330.00),
    ]

    def fill_sheet(ws, text_row=None):
        header(ws, ["Sale date", "Customer", "Amount"], [13, 26, 12])
        for r, (d, who, amt) in enumerate(rows, start=2):
            if r == text_row:
                ws.append([d.strftime("%m/%d/%Y"), who, amt])
                ws[f"A{r}"].number_format = "@"
            else:
                ws.append([d, who, amt])
                ws[f"A{r}"].number_format = US_DATE
            ws[f"C{r}"].number_format = USD
        ws.column_dimensions["E"].width = 34
        ws.column_dimensions["F"].width = 14
        ws["E1"], ws["F1"] = "Report date (type a date)", date(2026, 8, 31)
        ws["F1"].number_format = US_DATE
        ws["F1"].fill = PatternFill("solid", fgColor="FFF2CC")
        labels = [
            ("Year-to-date sales", '=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),1,1),A2:A500,"<="&F1)'),
            ("Month-to-date sales", '=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),MONTH(F1),1),A2:A500,"<="&F1)'),
            ("Whole report month (EOMONTH)", '=SUMIFS(C2:C500,A2:A500,">="&DATE(YEAR(F1),MONTH(F1),1),A2:A500,"<="&EOMONTH(F1,0))'),
            ("Start of year", "=DATE(YEAR(F1),1,1)"),
            ("Start of month", "=DATE(YEAR(F1),MONTH(F1),1)"),
            ("Month end", "=EOMONTH(F1,0)"),
            ("Rows with a real date (check)", "=COUNT(A2:A500)"),
            ("Rows with an amount (check)", "=COUNT(C2:C500)"),
        ]
        for i, (lab, f) in enumerate(labels, start=2):
            ws[f"E{i}"], ws[f"F{i}"] = lab, f
            ws[f"E{i}"].font = Font(bold=True)
            ws[f"F{i}"].fill = FORMULA_FILL
            ws[f"F{i}"].number_format = USD if i <= 4 else (US_DATE if i <= 7 else "0")
        ws["E1"].font = Font(bold=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    fill_sheet(ws)
    demo = wb.create_sheet("Text date demo")
    fill_sheet(demo, text_row=12)  # 08/18/2026 typed as text
    demo["E11"] = "Row 12 (08/18/2026) is stored as text on purpose."
    demo["E11"].font = Font(italic=True, color="C00000")

    notes_sheet(wb, [
        "Tidy Tabs: Year-to-Date Sales Total",
        "",
        "1. Sales tab: one row per sale. Type the Sale date (a real date), Customer, and Amount.",
        "2. F1 is the Report date. Type the date you want totals through. The sample uses 08/31/2026 so the numbers stay put; you can type =TODAY() instead for a live total.",
        "3. Year-to-date (F2) = SUMIFS(C2:C500, A2:A500, \">=\"&DATE(YEAR(F1),1,1), A2:A500, \"<=\"&F1). It adds amounts dated January 1 of the report year through the report date.",
        "4. Month-to-date (F3) is the same with DATE(YEAR(F1),MONTH(F1),1) as the start. F4 uses EOMONTH(F1,0) as the end to total the whole month.",
        "5. Change F1 and watch F2 to F4 move. Sales after the report date, like the 09/12/2026 and 10/03/2026 rows, are left out.",
        "6. Text date demo tab: the same data, but the 08/18/2026 date is stored as text. SUMIFS skips it; compare the totals with the Sales tab. F8 and F9 count real dates and amounts, so a gap means a text date.",
        "7. Add rows under the last sale. The formulas read down to row 500.",
        "",
        "Customers and amounts are fictional.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-year-to-date-sales-total.xlsx")


def quantity_price_tiers():
    wb = Workbook()
    tiers = wb.active
    tiers.title = "Tiers"
    header(tiers, ["Min quantity", "Unit price", "Tier"], [14, 12, 16])
    for r, (q, p, label) in enumerate([(1, 12.00, "1 to 11"), (12, 10.50, "12 to 35"),
                                        (36, 9.25, "36 to 71"), (72, 8.00, "72 and up")], start=2):
        tiers.append([q, p, label])
        tiers[f"B{r}"].number_format = USD
    tiers["E1"], tiers["E2"], tiers["E3"], tiers["E4"] = "Try a quantity", "Quantity", "Unit price", "Without IFERROR"
    tiers["F2"] = 0
    tiers["F3"] = '=IFERROR(VLOOKUP(F2,$A$2:$B$5,2,TRUE),"Below first tier")'
    tiers["F4"] = "=VLOOKUP(F2,$A$2:$B$5,2,TRUE)"
    tiers["F2"].fill = PatternFill("solid", fgColor="FFF2CC")
    for c in ("F3", "F4"):
        tiers[c].fill = FORMULA_FILL
        tiers[c].number_format = USD
    tiers["E1"].font = Font(bold=True)
    tiers.column_dimensions["E"].width = 18
    tiers.column_dimensions["F"].width = 18

    ws = wb.create_sheet("Orders")
    header(ws, ["Order date", "Customer", "Quantity", "VLOOKUP", "LOOKUP", "INDEX + MATCH",
                "Order total", "All 3 agree?"], [12, 24, 10, 12, 12, 15, 13, 13])
    orders = [
        (date(2026, 9, 2), "Harbor Coffee", 1),
        (date(2026, 9, 4), "Maple Street Bakery", 11),
        (date(2026, 9, 9), "Oak & Ember Candles", 12),
        (date(2026, 9, 14), "Lakeview Florist", 35),
        (date(2026, 9, 18), "Harbor Coffee", 36),
        (date(2026, 9, 23), "Maple Street Bakery", 71),
        (date(2026, 9, 26), "Oak & Ember Candles", 72),
        (date(2026, 9, 30), "Lakeview Florist", 500),
    ]
    for r, (d, who, qty) in enumerate(orders, start=2):
        ws.append([d, who, qty,
                   f"=VLOOKUP(C{r},Tiers!$A$2:$B$5,2,TRUE)",
                   f"=LOOKUP(C{r},Tiers!$A$2:$A$5,Tiers!$B$2:$B$5)",
                   f"=INDEX(Tiers!$B$2:$B$5,MATCH(C{r},Tiers!$A$2:$A$5,1))",
                   f"=C{r}*D{r}",
                   f'=IF(AND(D{r}=E{r},E{r}=F{r}),"Yes","CHECK")'])
        ws[f"A{r}"].number_format = US_DATE
        for col in "DEFG":
            ws[f"{col}{r}"].number_format = USD
        for col in "DEFGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    demo = wb.create_sheet("Unsorted demo")
    header(demo, ["Min quantity", "Unit price", "", "Quantity", "VLOOKUP result", "Correct price"],
           [14, 12, 3, 10, 16, 14])
    for r, (q, p) in enumerate([(36, 9.25), (1, 12.00), (72, 8.00), (12, 10.50)], start=2):
        demo.append([q, p])
        demo[f"B{r}"].number_format = USD
    for r, (qty, right) in enumerate([(5, 12.00), (20, 10.50), (40, 9.25), (80, 8.00)], start=2):
        demo[f"D{r}"] = qty
        demo[f"E{r}"] = f"=VLOOKUP(D{r},$A$2:$B$5,2,TRUE)"
        demo[f"F{r}"] = right
        demo[f"E{r}"].number_format = demo[f"F{r}"].number_format = USD
        demo[f"E{r}"].fill = FORMULA_FILL
    demo["A7"] = "The tier table above is NOT sorted from smallest to largest, so approximate match can return wrong prices."
    demo["A7"].font = Font(italic=True, color="C00000")

    notes_sheet(wb, [
        "Tidy Tabs: Wholesale Price Tiers by Quantity",
        "",
        "1. Tiers tab: list the smallest quantity of each tier in A and its unit price in B. Keep column A sorted from smallest to largest. Approximate match depends on that order.",
        "2. Orders tab: type the quantity in column C. The three unit price columns find the last tier whose minimum is less than or equal to the quantity.",
        "3. D: =VLOOKUP(C2,Tiers!$A$2:$B$5,2,TRUE). The TRUE (or leaving it out) turns on approximate match.",
        "4. E: =LOOKUP(C2,Tiers!$A$2:$A$5,Tiers!$B$2:$B$5). F: =INDEX(Tiers!$B$2:$B$5,MATCH(C2,Tiers!$A$2:$A$5,1)). The 1 in MATCH means the same approximate match.",
        "5. G: =C2*D2 is the order total. H checks that all three methods give the same price.",
        "6. Edge quantities in the sample: 11 pays the first tier, 12 pays the second; 35 and 36, 71 and 72 work the same way. Quantity 500 uses the last tier.",
        "7. A quantity below the first minimum (such as 0) returns #N/A. Tiers!F2 to F4 let you try one; F3 wraps the lookup in IFERROR.",
        "8. Unsorted demo tab: the same tiers in the wrong order. Compare the VLOOKUP result with the correct price. What an unsorted table returns is not guaranteed, so sort it.",
        "9. To add a tier, insert a row inside the table (not below it) so the ranges grow, and keep the minimums sorted.",
        "",
        "Customers, quantities, and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-quantity-price-tiers.xlsx")


def weighted_average_cost():
    wb = Workbook()
    ws = wb.active
    ws.title = "Purchases"
    header(ws, ["Purchase date", "Item", "Quantity", "Unit cost", "Line total"], [14, 20, 11, 11, 13])
    buys = [
        (date(2026, 7, 6), "Soy wax (lb)", 10, 4.00),
        (date(2026, 7, 20), "Cotton wicks", 200, 0.12),
        (date(2026, 8, 3), "Soy wax (lb)", 40, 3.50),
        (date(2026, 8, 12), "Glass jars", 48, 1.85),
        (date(2026, 8, 25), "Cotton wicks", 500, 0.09),
        (date(2026, 9, 7), "Soy wax (lb)", 25, 3.80),
        (date(2026, 9, 15), "Glass jars", 96, 1.60),
        (date(2026, 9, 22), "Cotton wicks", 300, 0.10),
        (date(2026, 9, 29), "Glass jars", 24, 2.10),
    ]
    for r, (d, item, qty, cost) in enumerate(buys, start=2):
        ws.append([d, item, qty, cost, f"=C{r}*D{r}"])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = '"$"#,##0.00'
        ws[f"E{r}"].number_format = USD
        ws[f"E{r}"].fill = FORMULA_FILL

    s = wb.create_sheet("Summary")
    header(s, ["Item", "Total quantity", "Total spent", "Weighted avg (SUMIFS)", "Weighted avg (SUMPRODUCT)",
               "Simple average", "Difference"], [20, 14, 13, 20, 24, 15, 12])
    for r, item in enumerate(["Soy wax (lb)", "Cotton wicks", "Glass jars"], start=2):
        s.append([item,
                  f"=SUMIFS(Purchases!$C$2:$C$40,Purchases!$B$2:$B$40,A{r})",
                  f"=SUMIFS(Purchases!$E$2:$E$40,Purchases!$B$2:$B$40,A{r})",
                  f"=C{r}/B{r}",
                  f"=SUMPRODUCT((Purchases!$B$2:$B$40=A{r})*Purchases!$C$2:$C$40*Purchases!$D$2:$D$40)/B{r}",
                  f"=AVERAGEIFS(Purchases!$D$2:$D$40,Purchases!$B$2:$B$40,A{r})",
                  f"=ROUND(D{r}-F{r},4)"])
        s[f"C{r}"].number_format = USD
        for col in "DEFG":
            s[f"{col}{r}"].number_format = '"$"#,##0.0000'
        for col in "BCDEFG":
            s[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Weighted Average Cost With SUMPRODUCT",
        "",
        "1. Purchases tab: one row per purchase with the date, item, quantity, and unit cost you paid. Column E is quantity times unit cost.",
        "2. Summary tab, column E: =SUMPRODUCT((Purchases!$B$2:$B$40=A2)*Purchases!$C$2:$C$40*Purchases!$D$2:$D$40)/B2. It multiplies quantity by unit cost for the rows that match the item, adds them up, and divides by the total quantity.",
        "3. Column D gives the same answer with SUMIFS: total spent divided by total quantity (=C2/B2).",
        "4. Column F, =AVERAGEIFS(...), is the simple average of the unit costs. It ignores how many units each purchase covered, so it is off whenever the quantities differ.",
        "5. Example: 10 lb at $4.00 and 40 lb at $3.50 cost $180.00 for 50 lb, so the weighted average is $3.60 per lb. The simple average of $4.00 and $3.50 is $3.75.",
        "6. The formulas read rows 2 to 40, so you can add purchases without editing them. Type each item name exactly the same way every time.",
        "7. This is an average unit cost of what you bought. It is not an inventory valuation method or tax advice.",
        "",
        "Items, quantities, and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-weighted-average-cost.xlsx")


def last_order_date_maxifs():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order date", "Customer", "Amount"], [13, 26, 12])
    orders = [
        (date(2026, 6, 8), "Harbor Coffee", 640.00),
        (date(2026, 6, 19), "Maple Street Bakery", 310.50),
        (date(2026, 7, 14), "Oak & Ember Candles", 455.00),
        (date(2026, 7, 28), "Harbor Coffee", 720.00),
        (date(2026, 8, 11), "Lakeview Florist", 180.25),
        (date(2026, 8, 26), "Maple Street Bakery", 395.00),
        (date(2026, 9, 9), "Oak & Ember Candles", 510.75),
        (date(2026, 9, 17), "Harbor Coffee", 865.00),
        (date(2026, 9, 24), "Lakeview Florist", 242.00),
        (date(2026, 9, 30), "Maple Street Bakery", 428.40),
    ]
    for r, (d, who, amt) in enumerate(orders, start=2):
        ws.append([d, who, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD

    s = wb.create_sheet("Summary")
    header(s, ["Customer", "Last order date", "Last order amount", "Days since last order", "Orders"],
           [26, 16, 18, 20, 9])
    for r, who in enumerate(["Harbor Coffee", "Maple Street Bakery", "Oak & Ember Candles", "Lakeview Florist"],
                            start=2):
        s.append([who,
                  f"=_xlfn.MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A{r})",
                  f"=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A{r},Orders!$A$2:$A$200,B{r})",
                  f"=$H$2-B{r}",
                  f"=COUNTIFS(Orders!$B$2:$B$200,A{r})"])
        s[f"B{r}"].number_format = US_DATE
        s[f"C{r}"].number_format = USD
        s[f"D{r}"].number_format = "0"
        for col in "BCDE":
            s[f"{col}{r}"].fill = FORMULA_FILL
    s.column_dimensions["G"].width = 14
    s.column_dimensions["H"].width = 14
    s["G2"], s["H2"] = "Report date", date(2026, 10, 9)
    s["H2"].number_format = US_DATE
    s["H2"].fill = PatternFill("solid", fgColor="FFF2CC")
    s["G2"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Last Order Date per Customer With MAXIFS",
        "",
        "1. Orders tab: one row per order with a real date, the customer name, and the amount. The Summary formulas read rows 2 to 200.",
        "2. Summary B: =MAXIFS(Orders!$A$2:$A$200,Orders!$B$2:$B$200,A2). It returns the latest (largest) date among the rows whose customer matches A2.",
        "3. Summary C: =SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2,Orders!$A$2:$A$200,B2). It adds the amounts for that customer on that last date. If a customer placed two orders on the same day, you get both added together.",
        "4. Summary D: report date minus the last order date. Change H2 (or type =TODAY()) to see it move.",
        "5. If a customer has no orders, MAXIFS returns 0, which shows as a date in 1899 or 1900. Wrap it in IF(COUNTIFS(...)=0,\"\",...) if your list can include new customers.",
        "6. MAXIFS is in Excel 2019 and Microsoft 365, and in Google Sheets. It is not in Excel 2016 or earlier; there, use an array formula with MAX and IF.",
        "7. Dates typed as text (left-aligned, no date format) are ignored by MAXIFS, so convert them to real dates first.",
        "",
        "Customers and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-last-order-date-maxifs.xlsx")


def count_orders_by_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order date", "Customer", "Amount"], [13, 26, 12])
    orders = [
        (date(2026, 7, 8), "Harbor Coffee", 120.00),
        (date(2026, 7, 22), "Maple Street Bakery", 245.50),
        (date(2026, 8, 5), "Oak & Ember Candles", 89.00),
        (date(2026, 8, 19), "Lakeview Florist", 310.25),
        (date(2026, 8, 31), "Harbor Coffee", 150.00),
        (date(2026, 9, 1), "Maple Street Bakery", 75.00),
        (date(2026, 9, 14), "Oak & Ember Candles", 420.00),
        (date(2026, 9, 28), "Lakeview Florist", 198.75),
        (date(2026, 10, 2), "Harbor Coffee", 260.00),
        (date(2026, 10, 7), "Maple Street Bakery", 135.50),
    ]
    for r, (d, who, amt) in enumerate(orders, start=2):
        ws.append([d, who, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD

    s = wb.create_sheet("Monthly")
    header(s, ["Month start", "Orders", "Revenue"], [14, 10, 14])
    for r, m in enumerate([date(2026, 7, 1), date(2026, 8, 1), date(2026, 9, 1), date(2026, 10, 1)], start=2):
        s.append([m,
                  f'=COUNTIFS(Orders!$A$2:$A$200,">="&A{r},Orders!$A$2:$A$200,"<"&EDATE(A{r},1))',
                  f'=SUMIFS(Orders!$C$2:$C$200,Orders!$A$2:$A$200,">="&A{r},Orders!$A$2:$A$200,"<"&EDATE(A{r},1))'])
        s[f"A{r}"].number_format = US_DATE
        s[f"C{r}"].number_format = USD
        s[f"B{r}"].fill = s[f"C{r}"].fill = FORMULA_FILL
    s["A6"], s["B6"], s["C6"] = "Total", "=SUM(B2:B5)", "=SUM(C2:C5)"
    s["A7"], s["B7"], s["C7"] = "Orders tab total", "=COUNT(Orders!A2:A200)", "=SUM(Orders!C2:C200)"
    s["C7"].number_format = s["C6"].number_format = USD
    for c in ("A6", "A7"):
        s[c].font = Font(bold=True)
    s.column_dimensions["A"].width = 18

    notes_sheet(wb, [
        "Tidy Tabs: Count Orders by Month With COUNTIFS",
        "",
        "1. Orders tab: one row per order with a real date, the customer, and the amount. The formulas read rows 2 to 200.",
        "2. Monthly tab, column A: type the first day of each month (07/01/2026). Format the cells as dates.",
        "3. Orders: =COUNTIFS(Orders!$A$2:$A$200,\">=\"&A2,Orders!$A$2:$A$200,\"<\"&EDATE(A2,1)). It counts dates on or after the first of the month and before the first of the next month.",
        "4. Revenue: the same two date tests inside SUMIFS, adding Orders!$C$2:$C$200.",
        "5. Why \"before the first of the next month\" and not \"on or before the last day\": it also works if a date carries a time of day, such as 08/31/2026 3:30 PM.",
        "6. Rows 6 and 7 compare the monthly totals with the whole Orders tab. If they differ, an order falls outside the months listed or its date is stored as text.",
        "7. Sample edge: 08/31/2026 counts in August and 09/01/2026 counts in September.",
        "",
        "Customers and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-count-orders-by-month.xlsx")


def average_order_value_by_channel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order date", "Channel", "Amount"], [13, 18, 12])
    orders = [
        (date(2026, 9, 3), "Etsy", 34.00),
        (date(2026, 9, 5), "Farmers market", 22.00),
        (date(2026, 9, 8), "Shopify", 78.00),
        (date(2026, 9, 12), "Etsy", 52.50),
        (date(2026, 9, 13), "Farmers market", 35.50),
        (date(2026, 9, 17), "Shopify", 45.00),
        (date(2026, 9, 20), "Farmers market", 18.50),
        (date(2026, 9, 24), "Etsy", 29.00),
        (date(2026, 9, 27), "Shopify", 96.50),
        (date(2026, 9, 28), "Farmers market", 40.00),
    ]
    for r, (d, ch, amt) in enumerate(orders, start=2):
        ws.append([d, ch, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD

    s = wb.create_sheet("By channel")
    header(s, ["Channel", "Orders", "Revenue", "Average order value", "Check (revenue / orders)"],
           [18, 10, 13, 20, 24])
    for r, ch in enumerate(["Etsy", "Shopify", "Farmers market"], start=2):
        s.append([ch,
                  f"=COUNTIFS(Orders!$B$2:$B$200,A{r})",
                  f"=SUMIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A{r})",
                  f"=AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A{r})",
                  f"=C{r}/B{r}"])
        for col in "CDE":
            s[f"{col}{r}"].number_format = USD
        for col in "BCDE":
            s[f"{col}{r}"].fill = FORMULA_FILL
    s["A5"], s["B5"], s["C5"], s["D5"] = "All channels", "=SUM(B2:B4)", "=SUM(C2:C4)", "=AVERAGE(Orders!C2:C200)"
    s["C5"].number_format = s["D5"].number_format = USD
    s["A5"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Average Order Value by Sales Channel With AVERAGEIFS",
        "",
        "1. Orders tab: one row per order with the date, the sales channel, and the order amount. Type each channel name exactly the same way every time.",
        "2. By channel, column D: =AVERAGEIFS(Orders!$C$2:$C$200,Orders!$B$2:$B$200,A2). The first argument is the range to average, then the range to test and the value it must match.",
        "3. Column B: =COUNTIFS(Orders!$B$2:$B$200,A2) counts the orders, and column C adds the revenue with SUMIFS.",
        "4. Column E divides revenue by orders. It should equal column D; if it does not, a row has a text amount or a blank.",
        "5. If a channel has no orders, AVERAGEIFS returns #DIV/0!. Wrap it in IFERROR(...,\"\") if you list channels you have not sold through yet.",
        "6. To average only one month, add two more date conditions, for example Orders!$A$2:$A$200,\">=\"&DATE(2026,9,1).",
        "7. Row 5 is the average across all channels. It is the average of orders, not the average of the three channel averages.",
        "",
        "Orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    wb.save(OUT / "tidy-tabs-average-order-value-by-channel.xlsx")


def letter_setup(ws, landscape=False, wrap_col_a=False):
    """US Letter, fit to one page wide, header row repeated."""
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    if wrap_col_a:
        for row in ws.iter_rows():
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
    else:
        ws.print_title_rows = "1:1"


def finish_letter(wb, landscape=False):
    for ws in wb.worksheets:
        letter_setup(ws, landscape=landscape and ws.title != "How to use",
                     wrap_col_a=ws.title == "How to use")


def subtotal_filtered_sales():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Sale date", "Channel", "Amount"], [13, 18, 12])
    sales = [
        (date(2026, 9, 2), "Shopify", 12.00),
        (date(2026, 9, 3), "Etsy", 34.00),
        (date(2026, 9, 5), "Farmers market", 22.00),
        (date(2026, 9, 6), "Farmers market", 8.50),
        (date(2026, 9, 8), "Shopify", 15.00),
        (date(2026, 9, 10), "Etsy", 18.50),
        (date(2026, 9, 12), "Farmers market", 6.00),
        (date(2026, 9, 14), "Shopify", 9.50),
        (date(2026, 9, 17), "Farmers market", 14.00),
        (date(2026, 9, 19), "Shopify", 11.00),
        (date(2026, 9, 21), "Etsy", 23.00),
        (date(2026, 9, 23), "Farmers market", 7.50),
        (date(2026, 9, 25), "Shopify", 5.00),
        (date(2026, 9, 27), "Farmers market", 10.50),
        (date(2026, 9, 29), "Shopify", 4.00),
    ]
    for r, (d, ch, amt) in enumerate(sales, start=2):
        ws.append([d, ch, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD
    ws.auto_filter.ref = "A1:C16"

    t = wb.create_sheet("Totals")
    header(t, ["Measure", "Formula", "Result"], [30, 28, 14])
    rows = [
        ("Visible total (SUBTOTAL 109)", "=SUBTOTAL(109,C2:C16)", "=SUBTOTAL(109,Sales!C2:C16)", USD),
        ("Visible count (SUBTOTAL 103)", "=SUBTOTAL(103,C2:C16)", "=SUBTOTAL(103,Sales!C2:C16)", "0"),
        ("Visible average (SUBTOTAL 101)", "=SUBTOTAL(101,C2:C16)", "=SUBTOTAL(101,Sales!C2:C16)", USD),
        ("All rows total (SUM, ignores filter)", "=SUM(C2:C16)", "=SUM(Sales!C2:C16)", USD),
        ("All rows count (COUNT)", "=COUNT(C2:C16)", "=COUNT(Sales!C2:C16)", "0"),
    ]
    for r, (label, shown, f, fmt) in enumerate(rows, start=2):
        t.append([label, None, f])
        t[f"B{r}"].value = shown
        t[f"B{r}"].data_type = "s"
        t[f"C{r}"].number_format = fmt
        t[f"C{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Sum Only the Visible Rows After Filtering",
        "",
        "1. Sales tab: one row per sale. The header row already has filter arrows (Data > Filter in Excel, Data > Create a filter in Google Sheets).",
        "2. Click the arrow on Channel and keep only one channel, for example Etsy. The other rows are hidden.",
        "3. Totals tab, C2: =SUBTOTAL(109,Sales!C2:C16) adds only the rows still visible. 109 means SUM, skipping rows hidden by a filter or by hand.",
        "4. C3: =SUBTOTAL(103,...) counts visible numbers. C4: =SUBTOTAL(101,...) averages the visible rows.",
        "5. C5: =SUM(Sales!C2:C16) adds every row, so it does not change when you filter. Compare it with C2.",
        "6. If you add rows below row 16, extend the ranges (or insert the new rows inside the list).",
        "7. Do not put SUBTOTAL cells inside the filtered range: they would be hidden or counted twice. That is why the totals sit on their own tab.",
        "",
        "Sales and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-subtotal-filtered-sales.xlsx")


def rolling_average_sales():
    wb = Workbook()
    ws = wb.active
    ws.title = "Monthly sales"
    header(ws, ["Month", "Sales", "3-month average", "Check (SUM / 3)", "Matches?"], [12, 12, 17, 17, 11])
    sales = [1200, 950, 1100, 2400, 3100, 1800, 1650, 1400, 2250, 2900, 3800, 4200]
    for r, amt in enumerate(sales, start=2):
        m = date(2026 if r < 2 + 12 else 2027, r - 1, 1)
        ws.append([m, amt])
        ws[f"A{r}"].number_format = "mmm yyyy"
        ws[f"B{r}"].number_format = USD
        if r >= 4:
            ws[f"C{r}"] = f'=IF(COUNT(B{r-2}:B{r})<3,"",AVERAGE(B{r-2}:B{r}))'
            ws[f"D{r}"] = f"=SUM(B{r-2}:B{r})/3"
            ws[f"E{r}"] = f'=IF(ROUND(C{r}-D{r},6)=0,"Yes","CHECK")'
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = USD

    notes_sheet(wb, [
        "Tidy Tabs: Rolling 3-Month Average of Sales",
        "",
        "1. Monthly sales tab: type each month's sales in column B, one row per month, oldest first. Months with no sales need a 0, not a blank.",
        "2. C4: =IF(COUNT(B2:B4)<3,\"\",AVERAGE(B2:B4)). It averages this month and the two before it. Fill it down. The first two months have no formula because they do not have three months yet.",
        "3. The COUNT test keeps the cell blank if any of the three months is empty or text, so a half-filled window is not averaged by mistake.",
        "4. D4: =SUM(B2:B4)/3 is a check. E says Yes when it equals column C.",
        "5. Each average moves one row at a time: row 5 drops January and adds April. That smooths out busy and slow seasons.",
        "6. To use 6 months instead of 3, widen the range (B2:B7) and change both the 3 in COUNT's test and the divisor.",
        "",
        "Months and sales are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-rolling-average-sales.xlsx")


def prorate_partial_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Proration"
    header(ws, ["Start date", "Monthly fee", "Days charged", "Days in month", "Prorated amount", "Example"],
           [13, 13, 13, 14, 16, 34])
    rows = [
        (date(2026, 10, 12), 150.00, "Retainer starts mid month"),
        (date(2026, 2, 20), 150.00, "February (28 days)"),
        (date(2026, 11, 1), 150.00, "Starts on the 1st: full month"),
        (date(2026, 12, 31), 150.00, "Starts on the last day: 1 day"),
        (date(2026, 9, 15), 85.00, "Farmers market booth, 30 day month"),
        (date(2026, 3, 10), 1200.00, "Studio rent"),
        (date(2028, 2, 20), 150.00, "Leap year February (29 days)"),
    ]
    for r, (d, fee, note) in enumerate(rows, start=2):
        ws.append([d, fee, f"=EOMONTH(A{r},0)-A{r}+1", f"=DAY(EOMONTH(A{r},0))", f"=ROUND(B{r}*C{r}/D{r},2)", note])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"B{r}"].number_format = USD
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = "0"
        ws[f"E{r}"].number_format = USD
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Prorate a Partial Month of Rent or a Retainer",
        "",
        "1. Proration tab: type the start date in A and the full monthly fee in B. Columns C to E are formulas; fill them down for more rows.",
        "2. C, days charged: =EOMONTH(A2,0)-A2+1. EOMONTH(A2,0) is the last day of the start month; the +1 counts the start day itself.",
        "3. D, days in that month: =DAY(EOMONTH(A2,0)). It is 28, 29, 30 or 31 depending on the month.",
        "4. E, prorated amount: =ROUND(B2*C2/D2,2), the fee times days charged over days in the month, rounded to cents.",
        "5. Example: $150.00 starting 10/12/2026 is 20 of 31 days, $96.77. $150.00 starting 02/20/2026 is 9 of 28 days, $48.21.",
        "6. Columns C and D must be formatted as Number. If they show a date like 01/20/1900, change the format to Number.",
        "7. This is arithmetic only. Some leases and contracts count a flat 30 day month or use other rules, so follow your own agreement. It is not legal advice.",
        "",
        "Dates and fees are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-prorate-partial-month.xlsx")


def two_way_rate_card():
    wb = Workbook()
    rc = wb.active
    rc.title = "Rate card"
    header(rc, ["Service", "Standard", "Rush", "Same day"], [16, 13, 13, 13])
    prices = [("Logo", 300, 420, 600), ("Brochure", 350, 500, 700), ("Website", 1200, 1600, 2100)]
    for r, (svc, *p) in enumerate(prices, start=2):
        rc.append([svc, *p])
        for col in "BCD":
            rc[f"{col}{r}"].number_format = USD

    q = wb.create_sheet("Quote")
    q.column_dimensions["A"].width = 34
    for col in "BCD":
        q.column_dimensions[col].width = 16
    q["A1"], q["A1"].font = "Pick a service and a turnaround (yellow cells)", Font(bold=True)
    q["A2"], q["B2"] = "Service", "Brochure"
    q["A3"], q["B3"] = "Turnaround", "Rush"
    for c in ("B2", "B3"):
        q[c].fill = PatternFill("solid", fgColor="FFF2CC")
    dv1 = DataValidation(type="list", formula1="='Rate card'!$A$2:$A$4", allow_blank=False)
    dv2 = DataValidation(type="list", formula1="='Rate card'!$B$1:$D$1", allow_blank=False)
    q.add_data_validation(dv1)
    q.add_data_validation(dv2)
    dv1.add("B2")
    dv2.add("B3")
    q["A5"] = "Price (INDEX + MATCH)"
    q["B5"] = "=INDEX('Rate card'!$B$2:$D$4,MATCH(B2,'Rate card'!$A$2:$A$4,0),MATCH(B3,'Rate card'!$B$1:$D$1,0))"
    q["A6"] = "Price (with IFERROR)"
    q["B6"] = ("=IFERROR(INDEX('Rate card'!$B$2:$D$4,MATCH(B2,'Rate card'!$A$2:$A$4,0),"
               "MATCH(B3,'Rate card'!$B$1:$D$1,0)),\"Check spelling\")")
    q["A7"] = "Price (SUMPRODUCT, comparison)"
    q["B7"] = "=SUMPRODUCT(('Rate card'!$A$2:$A$4=B2)*('Rate card'!$B$1:$D$1=B3)*'Rate card'!$B$2:$D$4)"
    for c in ("B5", "B6", "B7"):
        q[c].number_format = USD
        q[c].fill = FORMULA_FILL
        q[c].alignment = Alignment(horizontal="right")

    q["A9"] = "All 9 prices through the same formula"
    q["A9"].font = Font(bold=True)
    for i, col in enumerate("BCD"):
        q[f"{col}10"] = f"='Rate card'!{col}1"
        q[f"{col}10"].font = Font(bold=True)
    for r in (11, 12, 13):
        q[f"A{r}"] = f"='Rate card'!A{r-9}"
        q[f"A{r}"].font = Font(bold=True)
        for col in "BCD":
            q[f"{col}{r}"] = (f"=INDEX('Rate card'!$B$2:$D$4,MATCH($A{r},'Rate card'!$A$2:$A$4,0),"
                              f"MATCH({col}$10,'Rate card'!$B$1:$D$1,0))")
            q[f"{col}{r}"].number_format = USD
            q[f"{col}{r}"].fill = FORMULA_FILL
    q["A14"] = "Cells that match the rate card (should be 9)"
    q["B14"] = "=SUMPRODUCT(--(B11:D13='Rate card'!B2:D4))"

    notes_sheet(wb, [
        "Tidy Tabs: Two-Way Rate Card Lookup With INDEX and MATCH",
        "",
        "1. Rate card tab: services down column A, turnaround speeds across row 1, prices in B2:D4. Keep the labels spelled the same everywhere.",
        "2. Quote tab: pick a service in B2 and a turnaround in B3 from the dropdown lists.",
        "3. B5: =INDEX('Rate card'!$B$2:$D$4,MATCH(B2,'Rate card'!$A$2:$A$4,0),MATCH(B3,'Rate card'!$B$1:$D$1,0)). The first MATCH finds the row number, the second finds the column number, and INDEX returns the price where they cross.",
        "4. The 0 at the end of each MATCH means exact match.",
        "5. B6 wraps the same formula in IFERROR so a misspelled service or speed shows Check spelling instead of #N/A. The dropdowns stop most typos, but pasted text can still be wrong.",
        "6. B7 gets the same price with SUMPRODUCT. A misspelled label returns $0.00 there, not an error, so use it only as a cross-check.",
        "7. Rows 10 to 14 run the INDEX and MATCH formula on all 9 cells and count how many equal the rate card.",
        "8. To add a service or speed, insert a row or column inside the table and extend the ranges.",
        "",
        "Services and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-two-way-rate-card.xlsx")


def in_cell_bars():
    wb = Workbook()
    ws = wb.active
    ws.title = "Units sold"
    header(ws, ["Product", "Units sold", "Bar (block characters)", "Blocks (LEN)", "Bar (| fallback)"],
           [24, 12, 30, 13, 30])
    items = [("Soy candle, 8 oz", 3200), ("Beeswax melts", 1600), ("Gift tags (retired)", 0),
             ("Wick trimmer", 450), ("Lip balm", 2400), ("Mini jar candle", 800)]
    for r, (name, units) in enumerate(items, start=2):
        ws.append([name, units,
                   f'=REPT("█",ROUND(B{r}/MAX($B$2:$B$7)*20,0))',
                   f"=LEN(C{r})",
                   f'=REPT("|",ROUND(B{r}/MAX($B$2:$B$7)*20,0))'])
        ws[f"B{r}"].number_format = "#,##0"
        ws[f"C{r}"].font = Font(color="2E5E4E")
        ws[f"E{r}"].font = Font(color="2E5E4E")
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Bar Chart Inside Cells With REPT",
        "",
        "1. Units sold tab: product names in A, units in B (rows 2 to 7).",
        "2. C2: =REPT(\"█\",ROUND(B2/MAX($B$2:$B$7)*20,0)). REPT repeats the block character. The biggest value gets 20 blocks and every other row is scaled to it.",
        "3. Copy the █ character from the formula if you need to retype it. The $ signs keep the MAX range fixed when you fill down.",
        "4. D2: =LEN(C2) counts the blocks, as a check. With 3,200 / 1,600 / 0 / 450 units the bars are 20 / 10 / 0 / 3 blocks.",
        "5. Column E uses the | character instead. Use it if the block shows as an empty box or looks different in your font.",
        "6. To change the full bar length, replace the 20 in both the formula and your width. Widen the column so the longest bar fits.",
        "7. A bar is rounded to whole characters, so very small values can show 0 blocks.",
        "",
        "Products and units are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-in-cell-bars.xlsx")


def quarterly_sales_totals():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Sale date", "Amount", "Quarter"], [13, 13, 12])
    sales = [
        (date(2025, 12, 20), 90.00), (date(2026, 1, 15), 120.00), (date(2026, 2, 3), 85.50),
        (date(2026, 3, 31), 200.00), (date(2026, 4, 1), 150.00), (date(2026, 4, 18), 95.25),
        (date(2026, 5, 9), 310.00), (date(2026, 6, 30), 75.00), (date(2026, 7, 1), 220.00),
        (date(2026, 7, 22), 60.50), (date(2026, 8, 14), 180.00), (date(2026, 9, 30), 140.00),
        (date(2026, 10, 1), 99.99), (date(2026, 10, 5), 250.00), (date(2026, 11, 11), 130.00),
        (date(2026, 12, 31), 300.00),
    ]
    for r, (d, amt) in enumerate(sales, start=2):
        ws.append([d, amt, f'=YEAR(A{r})&" Q"&ROUNDUP(MONTH(A{r})/3,0)'])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"B{r}"].number_format = USD
        ws[f"C{r}"].fill = FORMULA_FILL

    q = wb.create_sheet("By quarter")
    header(q, ["Quarter", "Quarter starts", "Total (helper column)", "Orders", "Total (date range)", "Same?"],
           [12, 15, 21, 9, 19, 9])
    starts = [date(2025, 10, 1), date(2026, 1, 1), date(2026, 4, 1), date(2026, 7, 1), date(2026, 10, 1)]
    for r, d in enumerate(starts, start=2):
        q.append([f'=YEAR(B{r})&" Q"&ROUNDUP(MONTH(B{r})/3,0)', d,
                  f"=SUMIFS(Sales!$B$2:$B$200,Sales!$C$2:$C$200,A{r})",
                  f"=COUNTIFS(Sales!$C$2:$C$200,A{r})",
                  f'=SUMIFS(Sales!$B$2:$B$200,Sales!$A$2:$A$200,">="&B{r},Sales!$A$2:$A$200,"<"&EDATE(B{r},3))',
                  f'=IF(ROUND(C{r}-E{r},2)=0,"Yes","CHECK")'])
        q[f"B{r}"].number_format = US_DATE
        q[f"C{r}"].number_format = q[f"E{r}"].number_format = USD
        for col in "ACDEF":
            q[f"{col}{r}"].fill = FORMULA_FILL
    q["A7"], q["C7"], q["D7"], q["E7"] = "All quarters", "=SUM(C2:C6)", "=SUM(D2:D6)", "=SUM(E2:E6)"
    q["A8"], q["C8"] = "All sales on Sales tab", "=SUM(Sales!B2:B200)"
    q["D8"] = "=COUNT(Sales!B2:B200)"
    for c in ("C7", "E7", "C8"):
        q[c].number_format = USD
    for c in ("A7", "A8"):
        q[c].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Total Sales by Quarter",
        "",
        "1. Sales tab: type the sale date in A (a real date, not text) and the amount in B. Column C is a formula: fill it down.",
        "2. C2: =YEAR(A2)&\" Q\"&ROUNDUP(MONTH(A2)/3,0). ROUNDUP(MONTH/3,0) turns months 1 to 3 into 1, 4 to 6 into 2, 7 to 9 into 3 and 10 to 12 into 4. The year is added so 2025 Q4 and 2026 Q4 stay separate.",
        "3. By quarter tab, C2: =SUMIFS(Sales!$B$2:$B$200,Sales!$C$2:$C$200,A2) adds every sale with that quarter label. D counts the orders with COUNTIFS.",
        "4. E2 gets the same total without a helper column, by date range: =SUMIFS(amounts,dates,\">=\"&B2,dates,\"<\"&EDATE(B2,3)). B2 is the first day of the quarter. F says Yes when both agree.",
        "5. Row 7 adds the quarters and row 8 adds the whole Sales tab. If they differ, a date is outside the quarters listed or is stored as text.",
        "6. Quarters here are calendar quarters (January to March is Q1). If your business year starts in another month, this sheet does not apply as is.",
        "7. 03/31 and 04/01 fall in different quarters, and 12/31 is in Q4. Both methods handle the boundary days.",
        "",
        "Sales and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-quarterly-sales-totals.xlsx")


def missing_invoice_numbers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Invoice number", "Client", "Amount"], [16, 24, 13])
    rows = [(1001, "Maple Street Bakery", 450.00), (1002, "Dana Ruiz Design", 300.00),
            (1003, "Harbor Yoga", 180.00), (1005, "Cedar Hill Farm", 620.00),
            (1006, "Maple Street Bakery", 450.00), (1008, "Harbor Yoga", 180.00),
            (1009, "Sam Patel LLC", 975.00), (1009, "Sam Patel LLC", 975.00),
            (1011, "Dana Ruiz Design", 300.00), (1012, "Cedar Hill Farm", 620.00)]
    for r, (n, c, a) in enumerate(rows, start=2):
        ws.append([n, c, a])
        ws[f"C{r}"].number_format = USD

    ck = wb.create_sheet("Check")
    header(ck, ["Expected number", "Times found", "Status"], [17, 13, 14])
    for r in range(2, 14):
        ck.append([f"=MIN(Invoices!$A$2:$A$200)+ROW()-2",
                   f"=COUNTIF(Invoices!$A$2:$A$200,A{r})",
                   f'=IF(B{r}=0,"Missing",IF(B{r}>1,"Duplicate","OK"))'])
        ck[f"A{r}"].number_format = "0"
        for col in "ABC":
            ck[f"{col}{r}"].fill = FORMULA_FILL
    ck.conditional_formatting.add("C2:C13", FormulaRule(formula=['$C2="Missing"'], fill=ALERT_FILL))
    ck.conditional_formatting.add("C2:C13", FormulaRule(formula=['$C2="Duplicate"'], fill=PatternFill("solid", fgColor="FFF2CC")))
    ck["E1"], ck["F1"] = "Summary", None
    ck["E1"].font = Font(bold=True)
    ck.column_dimensions["E"].width = 28
    ck.column_dimensions["F"].width = 10
    summary = [
        ("Lowest number", "=MIN(Invoices!A2:A200)"),
        ("Highest number", "=MAX(Invoices!A2:A200)"),
        ("Numbers expected in range", "=F3-F2+1"),
        ("Invoices entered", "=COUNT(Invoices!A2:A200)"),
        ("Missing numbers", '=COUNTIF(C2:C13,"Missing")'),
        ("Numbers used twice", '=COUNTIF(C2:C13,"Duplicate")'),
        ("Check: entered - duplicate extras + missing", "=F5-(SUMPRODUCT((B2:B13>1)*(B2:B13-1)))+F6"),
    ]
    for i, (label, f) in enumerate(summary, start=2):
        ck[f"E{i}"], ck[f"F{i}"] = label, f
        ck[f"F{i}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Find Missing Invoice Numbers",
        "",
        "1. Invoices tab: list every invoice number you issued in column A (any order). Use plain numbers such as 1001. If your numbers have a prefix like INV-1001, keep the number in its own column.",
        "2. Check tab, A2: =MIN(Invoices!$A$2:$A$200)+ROW()-2 lists each expected number from the lowest one upward. Fill it down for as many numbers as the range needs.",
        "3. B2: =COUNTIF(Invoices!$A$2:$A$200,A2) counts how often that number appears.",
        "4. C2: =IF(B2=0,\"Missing\",IF(B2>1,\"Duplicate\",\"OK\")). Missing numbers turn red, duplicates yellow.",
        "5. F6 counts the missing numbers. F5 minus the extra copies plus F6 should equal F4 (the numbers expected in the range).",
        "6. The check only looks between your lowest and highest number. A missing number before the first or after the last one is not detected.",
        "7. A missing number is a question to answer, not proof of a problem: it may be a voided invoice. This sheet does not decide that.",
        "8. The Check tab lists 12 numbers. If your lowest-to-highest range is longer, fill the Check formulas down and widen the ranges in the summary; until then F8 will not equal F4, which tells you rows are missing from the list.",
        "",
        "Clients and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-missing-invoice-numbers.xlsx")


def sumif_contains_text():
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    header(ws, ["Date", "Description", "Amount"], [13, 34, 13])
    rows = [
        (date(2026, 9, 1), "Uber to client meeting", 18.40), (date(2026, 9, 2), "Printer ink cartridge", 42.99),
        (date(2026, 9, 4), "UBER EATS lunch", 14.25), (date(2026, 9, 7), "Pink ribbon spool", 6.50),
        (date(2026, 9, 9), "Etsy listing fees", 3.20), (date(2026, 9, 12), "Etsy shipping labels", 27.80),
        (date(2026, 9, 15), "Domain renewal", 12.00), (date(2026, 9, 18), "Uber airport", 36.10),
        (date(2026, 9, 21), "Packing tape and ink pen", 9.75), (date(2026, 9, 24), "Stamps*", 11.60),
        (date(2026, 9, 26), "Ink refill kit", 19.00), (date(2026, 9, 30), "Software subscription", 15.00),
    ]
    for r, (d, desc, amt) in enumerate(rows, start=2):
        ws.append([d, desc, amt])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD

    t = wb.create_sheet("Totals")
    header(t, ["Text to find", "Contains (total)", "Contains (count)", "Starts with (total)", "Ends with (total)"],
           [18, 17, 17, 19, 18])
    for r, word in enumerate(["Uber", "Etsy", "ink", "Ink pen", "Stamps~*", "Zoom"], start=2):
        t.append([word,
                  f'=SUMIF(Expenses!$B$2:$B$200,"*"&A{r}&"*",Expenses!$C$2:$C$200)',
                  f'=COUNTIF(Expenses!$B$2:$B$200,"*"&A{r}&"*")',
                  f'=SUMIF(Expenses!$B$2:$B$200,A{r}&"*",Expenses!$C$2:$C$200)',
                  f'=SUMIF(Expenses!$B$2:$B$200,"*"&A{r},Expenses!$C$2:$C$200)'])
        for col in "BDE":
            t[f"{col}{r}"].number_format = USD
        for col in "BCDE":
            t[f"{col}{r}"].fill = FORMULA_FILL
    t["A9"], t["B9"] = "All expenses", "=SUM(Expenses!C2:C200)"
    t["B9"].number_format = USD
    t["A9"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Sum Cells That Contain Certain Text",
        "",
        "1. Expenses tab: date, description and amount. Totals tab column A holds the word to look for.",
        "2. B2: =SUMIF(Expenses!$B$2:$B$200,\"*\"&A2&\"*\",Expenses!$C$2:$C$200). The * wildcard stands for any text, so the description only has to contain the word somewhere. Letter case is ignored.",
        "3. C2: =COUNTIF(Expenses!$B$2:$B$200,\"*\"&A2&\"*\") counts the matching rows.",
        "4. D2 drops the first * (text starts with the word). E2 drops the last * (text ends with the word).",
        "5. The match is on letters, not on words. \"ink\" also matches \"Pink ribbon spool\". Use a longer phrase such as \"Printer ink\" if that is a problem.",
        "6. To search for a real * or ?, put a ~ in front of it: Stamps~* finds the text Stamps*.",
        "7. A blank cell in column A makes the pattern ** , which matches every row that has text.",
        "8. SUMIF does not read cell colors and only matches text, so a number in the description column is not found by a * pattern.",
        "",
        "Expenses are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-sumif-contains-text.xlsx")


def invoice_installments():
    wb = Workbook()
    ws = wb.active
    ws.title = "Payment plan"
    ws.column_dimensions["A"].width = 24
    for col, w in zip("BCD", (16, 16, 16)):
        ws.column_dimensions[col].width = w
    ws["A1"], ws["A1"].font = "Invoice total", Font(bold=True)
    ws["B1"] = 2500.00
    ws["A2"], ws["A2"].font = "Number of payments (1 to 12)", Font(bold=True)
    ws["B2"] = 6
    ws["A3"], ws["A3"].font = "First payment date", Font(bold=True)
    ws["B3"] = date(2026, 11, 1)
    ws["B1"].number_format = USD
    ws["B3"].number_format = US_DATE
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=False,
                        showErrorMessage=True, errorTitle="Payments", error="Enter a whole number from 1 to 12.")
    ws.add_data_validation(dv)
    dv.add("B2")
    for c in ("B1", "B2", "B3"):
        ws[c].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["A4"] = "Regular payment"
    ws["B4"] = "=ROUND(B1/B2,2)"
    ws["A5"] = "Last payment"
    ws["B5"] = "=B1-B4*(B2-1)"
    for c in ("B4", "B5"):
        ws[c].number_format = USD
        ws[c].fill = FORMULA_FILL
    for col, name in zip("ABCD", ["Payment #", "Due date", "Amount", "Balance after"]):
        ws[f"{col}7"] = name
        ws[f"{col}7"].fill, ws[f"{col}7"].font = HEADER_FILL, HEADER_FONT
    for n in range(1, 13):
        r = 7 + n
        ws[f"A{r}"] = n
        ws[f"B{r}"] = f'=IF(A{r}>$B$2,"",EDATE($B$3,A{r}-1))'
        ws[f"C{r}"] = f'=IF(A{r}>$B$2,"",IF(A{r}=$B$2,$B$5,$B$4))'
        ws[f"D{r}"] = f'=IF(A{r}>$B$2,"",$B$1-SUM($C$8:C{r}))'
        ws[f"B{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = USD
        for col in "BCD":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["A21"], ws["C21"] = "Total of payments", "=SUM(C8:C19)"
    ws["A22"], ws["C22"] = "Difference from invoice", "=ROUND(C21-B1,2)"
    ws["C21"].number_format = ws["C22"].number_format = USD
    ws["A21"].font = ws["A22"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Split an Invoice Into Equal Monthly Payments",
        "",
        "1. Payment plan tab: change the three yellow cells: invoice total (B1), number of payments from 1 to 12 (B2) and the first payment date (B3).",
        "2. B4: =ROUND(B1/B2,2) is the regular payment rounded to cents.",
        "3. B5: =B1-B4*(B2-1) is the last payment. It takes whatever cents are left over so the payments add up to the invoice exactly. $1,000.00 in 3 payments is $333.33, $333.33, $333.34.",
        "4. Due dates: =IF(A8>$B$2,\"\",EDATE($B$3,A8-1)) moves the first date forward one month at a time. Rows beyond your number of payments stay blank.",
        "5. EDATE keeps the day of the month, but a month that is too short uses its last day: a first date of 01/31/2026 gives 02/28/2026 and then 03/31/2026.",
        "6. Balance after: =$B$1-SUM($C$8:C8), which should end at $0.00. C22 shows the difference between the payments and the invoice (should be $0.00).",
        "7. Only B2 values from 1 to 12 are supported. This is arithmetic for a plan you already agreed with the client; it adds no interest or fees.",
        "",
        "Amounts and dates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-invoice-installments.xlsx")


def marketplace_net_payout():
    wb = Workbook()
    ws = wb.active
    ws.title = "Net payout"
    header(ws, ["Order", "Sale price", "Fees", "Net payout", "Net % of price"], [26, 13, 12, 14, 15])
    orders = [("Soy candle, 8 oz", 18.00), ("Cedar soap bar", 8.50), ("Brass hoop earrings", 24.00),
              ("Gift tag set", 1.00), ("Candle gift box", 60.00), ("Sticker sample", 0.25)]
    for r, (name, price) in enumerate(orders, start=2):
        ws.append([name, price, f"=ROUND(B{r}*($H$2+$H$3)+$H$4,2)", f"=B{r}-C{r}", f"=D{r}/B{r}"])
        for col in "BCD":
            ws[f"{col}{r}"].number_format = USD
        ws[f"E{r}"].number_format = "0.0%"
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws["A8"], ws["B8"], ws["C8"], ws["D8"] = "Total", "=SUM(B2:B7)", "=SUM(C2:C7)", "=SUM(D2:D7)"
    ws["E8"] = "=D8/B8"
    for col in "BCD":
        ws[f"{col}8"].number_format = USD
    ws["E8"].number_format = "0.0%"
    ws["A8"].font = Font(bold=True)

    ws.column_dimensions["G"].width = 34
    ws.column_dimensions["H"].width = 12
    ws["G1"], ws["G1"].font = "Fee settings (fictional rates)", Font(bold=True)
    for r, (label, val, fmt) in enumerate([("Marketplace fee, % of price", 0.065, "0.0%"),
                                           ("Payment processing, % of price", 0.03, "0.0%"),
                                           ("Fixed fee per order", 0.25, USD)], start=2):
        ws[f"G{r}"], ws[f"H{r}"] = label, val
        ws[f"H{r}"].number_format = fmt
        ws[f"H{r}"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["G6"], ws["G6"].font = "Price needed for a target payout", Font(bold=True)
    ws["G7"], ws["H7"] = "Target net payout", 15.00
    ws["G8"], ws["H8"] = "Price to charge", "=ROUNDUP((H7+H4)/(1-H2-H3),2)"
    ws["G9"], ws["H9"] = "Check: net payout at that price", "=H8-ROUND(H8*(H2+H3)+H4,2)"
    ws["H7"].number_format = ws["H8"].number_format = ws["H9"].number_format = USD
    ws["H7"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["H8"].fill = ws["H9"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Net Payout After Marketplace Fees",
        "",
        "1. Net payout tab: yellow cells H2:H4 hold the fee settings. The rates here are made up. Replace them with the rates shown in your own seller account, and update them when the marketplace changes its fees.",
        "2. C2: =ROUND(B2*($H$2+$H$3)+$H$4,2) is the fee on that order: both percentages of the price plus the fixed fee, rounded to cents.",
        "3. D2: =B2-C2 is the net payout. E2: =D2/B2 is the share of the price you keep.",
        "4. The fixed fee hurts small orders most: the $1.00 gift tag set keeps only 65.0% of its price, and the $0.25 sticker sample ends up at a net payout of -$0.02, a loss.",
        "5. H8: =ROUNDUP((H7+H4)/(1-H2-H3),2) works backward to the price that nets your target, rounded up to the next cent. H9 checks it by running the forward formula on that price. Because fees are rounded to cents, the check can land a cent above the target (here $15.01 for a $15.00 target).",
        "6. This sheet covers only the fees you type in. It does not include shipping, sales tax collected, advertising, refunds, or income tax.",
        "",
        "Orders and rates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-marketplace-net-payout.xlsx")


INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")


def invoice_due_dates():
    wb = Workbook()
    ws = wb.active
    ws.title = "Due dates"
    header(ws, ["Invoice", "Client", "Invoice date", "Terms", "Days", "Due date", "Days until due", "Status"],
           [12, 26, 13, 17, 14, 13, 14, 12])
    invoices = [
        ("INV-1001", "Maple Street Bakery", date(2026, 9, 1), "Net 30"),
        ("INV-1002", "Harbor Yoga Studio", date(2026, 9, 15), "Net 15"),
        ("INV-1003", "Birch Design Co", date(2026, 9, 28), "Net 45"),
        ("INV-1004", "Lantern Print Shop", date(2026, 10, 5), "Net 30"),
        ("INV-1005", "Cedar Row Realty", date(2026, 10, 13), "Due on receipt"),
        ("INV-1006", "Willow Pet Care", date(2026, 10, 1), "Net 60"),
        ("INV-1007", "Granite Fitness", date(2026, 8, 20), "Net 60"),
        ("INV-1008", "Juniper Books", date(2026, 10, 10), "Net 15"),
    ]
    last = 20
    for r in range(2, last + 1):
        if r - 2 < len(invoices):
            ws.append(list(invoices[r - 2]))
        else:
            ws.append([])
        ws[f"C{r}"].number_format = US_DATE
        ws[f"E{r}"] = f'=IF(D{r}="","",IFERROR(VLOOKUP(D{r},$J$2:$K$6,2,FALSE),"Check terms"))'
        ws[f"F{r}"] = f'=IF(OR(C{r}="",NOT(ISNUMBER(E{r}))),"",C{r}+E{r})'
        ws[f"G{r}"] = f'=IF(F{r}="","",F{r}-$K$8)'
        ws[f"H{r}"] = f'=IF(G{r}="","",IF(G{r}<0,"Past due",IF(G{r}=0,"Due today","Open")))'
        ws[f"F{r}"].number_format = US_DATE
        for col in "EFGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1="=$J$2:$J$6", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"D2:D{last}")
    ws.conditional_formatting.add(f"H2:H{last}", FormulaRule(formula=['H2="Past due"'], fill=ALERT_FILL))
    ws.column_dimensions["I"].width = 4
    ws.column_dimensions["J"].width = 20
    ws.column_dimensions["K"].width = 12
    ws["J1"], ws["K1"] = "Terms", "Days"
    for c in (ws["J1"], ws["K1"]):
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for r, (name, days) in enumerate([("Due on receipt", 0), ("Net 15", 15), ("Net 30", 30),
                                      ("Net 45", 45), ("Net 60", 60)], start=2):
        ws[f"J{r}"], ws[f"K{r}"] = name, days
        ws[f"J{r}"].fill = ws[f"K{r}"].fill = INPUT_FILL
    ws["J8"], ws["K8"] = "As-of date", date(2026, 10, 13)
    ws["K8"].number_format = US_DATE
    ws["J8"].font = Font(bold=True)
    ws["K8"].fill = INPUT_FILL
    ws["J10"], ws["K10"] = "Open or due today", f'=COUNTIF(H2:H{last},"Open")+COUNTIF(H2:H{last},"Due today")'
    ws["J11"], ws["K11"] = "Past due", f'=COUNTIF(H2:H{last},"Past due")'

    notes_sheet(wb, [
        "Tidy Tabs: Calculate an Invoice Due Date From Payment Terms",
        "",
        "1. Due dates tab: type the invoice number, client, invoice date (C) and pick the terms from the dropdown in column D. Rows 2 to 20 have the formulas ready.",
        "2. The terms table sits in J2:K6 (yellow). Each name has a number of days. Add or change terms by editing those cells; keep the names exactly as the dropdown shows them.",
        "3. E2: =IF(D2=\"\",\"\",IFERROR(VLOOKUP(D2,$J$2:$K$6,2,FALSE),\"Check terms\")) looks up the days. A term that is not in the table shows Check terms.",
        "4. F2: =IF(OR(C2=\"\",NOT(ISNUMBER(E2))),\"\",C2+E2) is the due date: the invoice date plus the days. Dates are plain numbers, so adding days works across month ends (08/20/2026 plus 60 days is 10/19/2026).",
        "5. The short version of the whole sheet is =C2+VLOOKUP(D2,$J$2:$K$6,2,FALSE).",
        "6. K8 is the as-of date. Type =TODAY() there to compare with the current day. G2 is the due date minus the as-of date (negative means days late). H2 says Past due, Due today or Open.",
        "7. The due date is not moved for weekends or holidays, and it does not read terms such as 2/10 net 30 or end of month. Use WORKDAY if you want business days.",
        "8. This is date arithmetic only. The terms on a real invoice come from your contract with the client.",
        "",
        "Invoices, clients and dates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-invoice-due-dates.xlsx")


def sales_by_weekday():
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily sales"
    header(ws, ["Date", "Sales", "Weekday"], [13, 12, 11])
    amounts = [42.00, 38.50, 51.25, 65.00, 148.00, 212.50, 96.00,
               40.00, 35.00, 47.50, 70.25, 155.00, 230.00, 88.50,
               44.50, 41.00, 55.00, 62.00, 162.75, 248.00, 101.00,
               39.00, 36.25, 49.00, 68.00, 171.00, 226.50, 93.00]
    last = 60
    for i in range(2, last + 1):
        if i - 2 < len(amounts):
            ws.append([date(2026, 9, 1 + i - 2), amounts[i - 2]])
        else:
            ws.append([])
        ws[f"A{i}"].number_format = US_DATE
        ws[f"B{i}"].number_format = USD
        ws[f"C{i}"] = f'=IF(A{i}="","",TEXT(A{i},"ddd"))'
        ws[f"C{i}"].fill = FORMULA_FILL

    s = wb.create_sheet("By weekday")
    header(s, ["Weekday", "Days counted", "Total sales", "Average per day", "Share of sales", "Best day"],
           [14, 14, 14, 17, 15, 12])
    rng = "'Daily sales'!$C$2:$C$200"
    for r, day in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], start=2):
        s.append([day, f"=COUNTIF({rng},A{r})",
                  f"=SUMIF({rng},A{r},'Daily sales'!$B$2:$B$200)",
                  f'=IF(B{r}=0,"",C{r}/B{r})',
                  f'=IF($C$9=0,"",C{r}/$C$9)',
                  f'=IF(D{r}="","",IF(D{r}=MAX($D$2:$D$8),"Best day",""))'])
        s[f"C{r}"].number_format = s[f"D{r}"].number_format = USD
        s[f"E{r}"].number_format = "0.0%"
        for col in "BCDEF":
            s[f"{col}{r}"].fill = FORMULA_FILL
    s["A9"], s["B9"], s["C9"] = "Total", "=SUM(B2:B8)", "=SUM(C2:C8)"
    s["A9"].font = Font(bold=True)
    s["C9"].number_format = USD
    s["A11"] = "Check: days counted match the Daily sales tab"
    s["F11"] = "=IF(B9=COUNT('Daily sales'!A2:A200),\"OK\",\"CHECK\")"
    s["A12"] = "Check: total matches the Daily sales tab"
    s["F12"] = "=IF(ROUND(C9-SUM('Daily sales'!B2:B200),2)=0,\"OK\",\"CHECK\")"

    notes_sheet(wb, [
        "Tidy Tabs: Total Sales by Day of the Week",
        "",
        "1. Daily sales tab: one row per day, with the date in A and that day's sales in B. Column C names the weekday.",
        "2. C2: =IF(A2=\"\",\"\",TEXT(A2,\"ddd\")) turns a date into Mon, Tue and so on. Fill it down as far as your data goes (it is ready to row 60).",
        "3. By weekday tab, C2: =SUMIF('Daily sales'!$C$2:$C$200,A2,'Daily sales'!$B$2:$B$200) adds the sales for each weekday. B2 counts the days with COUNTIF.",
        "4. D2: =IF(B2=0,\"\",C2/B2) is the average per day. The short form is =AVERAGEIF('Daily sales'!$C$2:$C$200,\"Sat\",'Daily sales'!$B$2:$B$200).",
        "5. Compare averages, not totals, when some weekdays appear more often in your data. F2 marks the weekday with the highest average.",
        "6. Dates should be real dates. A date stored as text may still get a weekday name in column C, but the COUNT of real dates skips it, so check cell F11 shows CHECK. Re-enter it as a date.",
        "7. A day with $0.00 sales still counts as a day if it has a row. Days with no row are not counted.",
        "8. The weekday names are English (Mon to Sun). They work the same for a US-English file in Excel and Google Sheets.",
        "",
        "Dates and sales are fictional (28 days from 09/01/2026). Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-sales-by-weekday.xlsx")


def repeat_customers():
    wb = Workbook()
    o = wb.active
    o.title = "Orders"
    header(o, ["Order", "Customer", "Order date", "Total"], [10, 20, 13, 12])
    orders = [
        ("1001", "Ana Ruiz", date(2026, 8, 4), 42.00), ("1002", "Eli Brooks", date(2026, 8, 6), 18.00),
        ("1003", "Ben Ortiz", date(2026, 8, 12), 65.50), ("1004", "Ana Ruiz", date(2026, 8, 20), 30.00),
        ("1005", "Cara Lee", date(2026, 8, 25), 24.00), ("1006", "Eli Brooks", date(2026, 9, 2), 36.00),
        ("1007", "Dev Patel", date(2026, 9, 9), 54.25), ("1008", "Fay Nolan", date(2026, 9, 14), 12.00),
        ("1009", "Eli Brooks", date(2026, 9, 21), 28.75), ("1010", "Cara Lee", date(2026, 9, 28), 48.00),
        ("1011", "Hana Cole", date(2026, 10, 1), 90.00), ("1012", "Ana Ruiz", date(2026, 10, 5), 22.50),
        ("1013", "Eli Brooks", date(2026, 10, 9), 31.00),
    ]
    for r, row in enumerate(orders, start=2):
        o.append(list(row))
        o[f"C{r}"].number_format = US_DATE
        o[f"D{r}"].number_format = USD

    c = wb.create_sheet("Customers")
    header(c, ["Customer", "Orders", "Type", "Total spent"], [20, 10, 12, 14])
    names = ["Ana Ruiz", "Ben Ortiz", "Cara Lee", "Dev Patel", "Eli Brooks", "Fay Nolan", "Gus Hall", "Hana Cole"]
    last = 31
    for r in range(2, last + 1):
        c.append([names[r - 2]] if r - 2 < len(names) else [])
        c[f"B{r}"] = f'=IF(A{r}="","",COUNTIF(Orders!$B$2:$B$200,A{r}))'
        c[f"C{r}"] = f'=IF(A{r}="","",IF(B{r}=0,"No orders",IF(B{r}=1,"One time","Repeat")))'
        c[f"D{r}"] = f'=IF(A{r}="","",SUMIF(Orders!$B$2:$B$200,A{r},Orders!$D$2:$D$200))'
        c[f"D{r}"].number_format = USD
        for col in "BCD":
            c[f"{col}{r}"].fill = FORMULA_FILL
    c.column_dimensions["E"].width = 4
    c.column_dimensions["F"].width = 36
    c.column_dimensions["G"].width = 14
    c["F1"] = "Summary"
    c["F1"].font = Font(bold=True)
    summary = [
        ("Customers with at least one order", f'=COUNTIF(B2:B{last},">0")', "0"),
        ("Repeat customers (2 or more orders)", f'=COUNTIF(B2:B{last},">=2")', "0"),
        ("Repeat customer rate", '=IF(G2=0,"",G3/G2)', "0.0%"),
        ("Orders counted above", f"=SUM(B2:B{last})", "0"),
        ("Orders on the Orders tab", "=COUNTA(Orders!B2:B200)", "0"),
        ("Check: every order matched a customer", '=IF(G5=G6,"OK","CHECK")', "General"),
        ("Spent by repeat customers", f'=SUMIF(C2:C{last},"Repeat",D2:D{last})', USD),
        ("Share of spending from repeat customers", f'=IF(SUM(D2:D{last})=0,"",G8/SUM(D2:D{last}))', "0.0%"),
    ]
    for r, (label, f, fmt) in enumerate(summary, start=2):
        c[f"F{r}"], c[f"G{r}"] = label, f
        c[f"G{r}"].number_format = fmt
        c[f"G{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Repeat Customer Rate With COUNTIF",
        "",
        "1. Orders tab: one row per order, with the customer name in column B. Paste or type your orders under the sample rows.",
        "2. Customers tab: list each customer once in column A . Rows 2 to 31 have formulas ready.",
        "3. B2: =IF(A2=\"\",\"\",COUNTIF(Orders!$B$2:$B$200,A2)) counts that customer's orders. C2 labels them No orders, One time or Repeat. D2 adds up what they spent with SUMIF.",
        "4. G2: =COUNTIF(B2:B31,\">0\") counts customers who ordered at all. G3: =COUNTIF(B2:B31,\">=2\") counts repeat customers.",
        "5. G4: =IF(G2=0,\"\",G3/G2) is the repeat customer rate. The short form is =COUNTIF(B2:B9,\">=2\")/COUNTIF(B2:B9,\">0\").",
        "6. G5 to G7 check that every order on the Orders tab matched a name on the Customers tab. A misspelled name or an extra space at the end of a cell makes G7 show CHECK.",
        "7. COUNTIF ignores capital letters, so ana ruiz and Ana Ruiz are the same customer. It does not ignore spaces or different spellings.",
        "8. The rate covers only the orders you pasted in. Choose the same date range for every rate you compare.",
        "",
        "Names, orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-repeat-customers.xlsx")


def ounces_to_pounds_ounces():
    wb = Workbook()
    ws = wb.active
    ws.title = "Weights"
    header(ws, ["Item", "Packed ounces", "Pounds", "Leftover ounces", "Label"], [26, 14, 9, 16, 16])
    items = [("Lavender candle box", 50), ("Oat soap bundle", 16), ("Earring card", 3.5), ("Canvas tote bag", 22),
             ("Ceramic mug pair", 41), ("Wax melt case", 96), ("Greeting card set", 12.25), ("Tea towel", 9),
             ("Gift crate", 135), ("Print tube", 28)]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(items[r - 2]) if r - 2 < len(items) else [])
        ws[f"B{r}"].fill = INPUT_FILL
        ws[f"C{r}"] = f'=IF(B{r}="","",INT(B{r}/16))'
        ws[f"D{r}"] = f'=IF(B{r}="","",MOD(B{r},16))'
        ws[f"E{r}"] = f'=IF(B{r}="","",INT(B{r}/16)&" lb "&MOD(B{r},16)&" oz")'
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"D{r}"].number_format = "0.00"
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 34
    ws.column_dimensions["H"].width = 16
    ws["G1"] = "Whole shipment"
    ws["G1"].font = Font(bold=True)
    summary = [
        ("Packages with a weight", f"=COUNT(B2:B{last})", "0"),
        ("Total ounces", f"=ROUND(SUM(B2:B{last}),2)", "0.00"),
        ("Total pounds", "=INT(H3/16)", "0"),
        ("Leftover ounces", "=ROUND(MOD(H3,16),2)", "0.00"),
        ("Total label", '=H4&" lb "&H5&" oz"', "General"),
        ("Check: pounds and ounces add back up", f'=IF(ROUND(SUM(C2:C{last})*16+SUM(D2:D{last})-H3,6)=0,"OK","CHECK")', "General"),
    ]
    for r, (label, f, fmt) in enumerate(summary, start=2):
        ws[f"G{r}"], ws[f"H{r}"] = label, f
        ws[f"H{r}"].number_format = fmt
        ws[f"H{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Ounces to Pounds and Ounces With INT and MOD",
        "",
        "1. Weights tab: type each item and its packed weight in ounces. Yellow cells are yours (B2:B30). Gray cells are formulas.",
        "2. C2: =IF(B2=\"\",\"\",INT(B2/16)) gives the whole pounds. 16 ounces make 1 pound.",
        "3. D2: =IF(B2=\"\",\"\",MOD(B2,16)) gives the ounces left over after the whole pounds.",
        "4. E2: =IF(B2=\"\",\"\",INT(B2/16)&\" lb \"&MOD(B2,16)&\" oz\") joins them into text such as 3 lb 2 oz.",
        "5. H3 to H6 do the same for the total of all packages (rounded to 2 decimals so totals do not show long decimals). H7 turns the pounds and ounces back into ounces and shows CHECK if they do not match the total.",
        "6. The label in column E is text. Use columns C and D when you need numbers for a label printer or a carrier form.",
        "7. A blank weight stays blank. A weight of 0 shows 0 lb 0 oz.",
        "8. This only converts units. It does not look up a carrier rate or a weight limit. Read those on your carrier's own pages.",
        "",
        "Items and weights are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-ounces-to-pounds-ounces.xlsx")


def listing_length_check():
    wb = Workbook()
    ws = wb.active
    ws.title = "Listings"
    header(ws, ["Title", "Tags (separated by commas)", "Characters", "Tags", "Title check", "Tag check"],
           [50, 52, 11, 8, 13, 12])
    listings = [
        ("Handmade Lavender Soy Candle 8 oz Gift for Her", "lavender candle, soy candle, gift for her, 8 oz candle"),
        ("Cedar and Sage Soy Candle Small Batch 8 oz Hand Poured Gift", "cedar candle, sage candle, hand poured, soy wax, small batch, housewarming, gift, 8 oz"),
        ("Oatmeal Goat Milk Soap Bar Handmade", "oatmeal soap, goat milk soap, handmade soap, gentle soap, bath gift, soap bar, natural soap, sensitive skin, farmers market"),
        ("Brass Hoop Earrings Minimalist Gold Tone Everyday Jewelry for Women Gift Idea", "brass hoops, gold hoop earrings, minimalist, everyday earrings"),
        ("Letterpress Birthday Card Blank Inside", "birthday card,letterpress,blank card"),
        ("Charcoal Soap Bar Cold Process Vegan", ""),
        ("Eucalyptus Mint Bath Salts 12 oz Jar", "bath salts, eucalyptus, spa gift,"),
        ("Beeswax Taper Candles Pair Natural Unscented", "taper candles, beeswax, natural candle, dinner candles, unscented"),
        ("Linen Wick Soy Candle Tin Travel Size 4 oz", "travel candle, tin candle, soy candle, stocking stuffer, 4 oz, gift, small candle, linen"),
        ("Kraft Gift Box Set Soap and Candle Combo Birthday", "gift set, soap and candle, birthday gift, kraft box, self care"),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(listings[r - 2]) if r - 2 < len(listings) else [])
        ws[f"C{r}"] = f'=IF(A{r}="","",LEN(A{r}))'
        ws[f"D{r}"] = (f'=IF(AND(A{r}="",B{r}=""),"",IF(TRIM(B{r})="",0,LEN(B{r})-LEN(SUBSTITUTE(B{r},",",""))+1'
                       f'-(RIGHT(TRIM(B{r}),1)=",")))')
        ws[f"E{r}"] = f'=IF(A{r}="","",IF(C{r}>$I$2,"Too long","OK"))'
        ws[f"F{r}"] = f'=IF(AND(A{r}="",B{r}=""),"",IF(D{r}>$I$3,"Too many","OK"))'
        for col in "CDEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"E2:F{last}", FormulaRule(formula=['OR(E2="Too long",E2="Too many")'], fill=ALERT_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 36
    ws.column_dimensions["I"].width = 12
    ws["H1"] = "Limits and summary"
    ws["H1"].font = Font(bold=True)
    rows = [
        ("Title limit (characters), type yours", 60, "0"),
        ("Tag limit, type yours", 8, "0"),
        ("Listings checked", f'=COUNTA(A2:A{last})', "0"),
        ("Titles too long", f'=COUNTIF(E2:E{last},"Too long")', "0"),
        ("Listings with too many tags", f'=COUNTIF(F2:F{last},"Too many")', "0"),
        ("Check: characters add up", f'=IF(SUM(C2:C{last})=SUMPRODUCT(LEN(A2:A{last})),"OK","CHECK")', "General"),
    ]
    for r, (label, v, fmt) in enumerate(rows, start=2):
        ws[f"H{r}"], ws[f"I{r}"] = label, v
        ws[f"I{r}"].number_format = fmt
        ws[f"I{r}"].fill = INPUT_FILL if r in (2, 3) else FORMULA_FILL
    ws["H9"] = "Sample limits only. Copy the real"
    ws["H10"] = "numbers from your marketplace's help page."

    notes_sheet(wb, [
        "Tidy Tabs: Count Characters and Tags in a Product Listing",
        "",
        "1. Listings tab: type or paste each title in column A and its tags in column B, separated by commas. Rows 2 to 30 have formulas ready.",
        "2. I2 and I3 are yellow input cells. Type the title limit and the tag limit from your marketplace's own help page. The 60 and 8 in the sample are made-up numbers, not any marketplace's rules.",
        "3. C2: =IF(A2=\"\",\"\",LEN(A2)) counts every character in the title, including spaces.",
        "4. D2 counts the tags: =LEN(B2)-LEN(SUBSTITUTE(B2,\",\",\"\"))+1 counts the commas and adds 1. The template wraps it so an empty tags cell gives 0 and a comma left at the end is not counted as a tag.",
        "5. E and F show OK, Too long or Too many. A title exactly at the limit is OK.",
        "6. I7 adds up the character counts a second way. If it shows CHECK, a formula in column C was overwritten.",
        "7. LEN counts spaces at the start and end of a cell. Spaces after commas do not change the tag count, but two commas in a row do (an empty tag is still counted).",
        "",
        "Listings are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-listing-length-check.xlsx")


def supplier_price_compare():
    wb = Workbook()
    ws = wb.active
    ws.title = "Compare"
    header(ws, ["Supplier", "Package price", "Package size (oz)", "Shipping", "Cost per oz",
                "Cost for amount needed", "Rank"], [24, 14, 14, 11, 13, 17, 8])
    quotes = [
        ("Pine Ridge Wax Co", 42.50, 160, 12.00),
        ("Lakeside Supply", 27.99, 80, 9.50),
        ("Harbor Craft Wholesale", 64.00, 240, 0.00),
        ("Birchwood Candle Goods", 18.50, 48, 7.25),
        ("Maple Lane Supply", 33.00, 100, 8.00),
    ]
    for r, q in enumerate(quotes, start=2):
        ws.append(list(q))
        ws[f"E{r}"] = f'=IF(OR(B{r}="",C{r}<=0),"",(B{r}+D{r})/C{r})'
        ws[f"F{r}"] = f'=IF(E{r}="","",ROUND(E{r}*$I$2,2))'
        ws[f"G{r}"] = f'=IF(E{r}="","",RANK(E{r},$E$2:$E$6,1))'
        ws[f"B{r}"].number_format = ws[f"D{r}"].number_format = ws[f"F{r}"].number_format = USD
        ws[f"E{r}"].number_format = '"$"0.0000'
        for col in "EFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        for col in "ABCD":
            ws[f"{col}{r}"].fill = INPUT_FILL
    ws.column_dimensions["H"].width = 36
    ws.column_dimensions["I"].width = 20
    ws["H1"] = "Compare"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["I2"] = "Amount you need (oz)", 200
    ws["I2"].fill = INPUT_FILL
    ws["H4"] = "Cheapest supplier per oz"
    ws["I4"] = '=IFERROR(INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0)),"Enter prices")'
    ws["H5"], ws["I5"] = "Lowest cost per oz", '=IF(COUNT(E2:E6)=0,"",MIN(E2:E6))'
    ws["H6"], ws["I6"] = "Highest cost per oz", '=IF(COUNT(E2:E6)=0,"",MAX(E2:E6))'
    ws["H7"], ws["I7"] = "Cost of amount needed, cheapest", '=IF(COUNT(F2:F6)=0,"",MIN(F2:F6))'
    ws["H8"], ws["I8"] = "Extra paid with the priciest", '=IF(COUNT(F2:F6)=0,"",MAX(F2:F6)-MIN(F2:F6))'
    for a in ("I5", "I6"):
        ws[a].number_format = '"$"0.0000'
    for a in ("I7", "I8"):
        ws[a].number_format = USD
    for a in ("I4", "I5", "I6", "I7", "I8"):
        ws[a].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Find the Cheapest Supplier by Cost Per Ounce",
        "",
        "1. Compare tab: type each supplier's package price (B), package size in ounces (C) and shipping for that package (D). The yellow cells are inputs.",
        "2. E2: =IF(OR(B2=\"\",C2<=0),\"\",(B2+D2)/C2) is the cost per ounce including shipping. A blank row or a size of 0 stays blank.",
        "3. I4: =INDEX($A$2:$A$6,MATCH(MIN($E$2:$E$6),$E$2:$E$6,0)) returns the name of the supplier with the lowest cost per ounce. MIN finds the lowest number, MATCH finds its row, INDEX reads the name.",
        "4. If two suppliers tie, the one listed first is returned. Column G (RANK) shows the tie.",
        "5. I2 is how many ounces you need. F2 multiplies each supplier's cost per ounce by it, rounded to cents.",
        "6. Use one unit for every row (all ounces, or all pounds). Mixing units gives a wrong answer without any error.",
        "7. Shipping is counted once per package. If you buy several packages, or shipping is free above an order size, adjust the shipping cell yourself. Minimum order sizes and delivery times are not compared.",
        "",
        "Suppliers and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-supplier-price-compare.xlsx")


def split_sku_code():
    wb = Workbook()
    ws = wb.active
    ws.title = "SKUs"
    header(ws, ["SKU", "Units sold", "Category", "Style", "Size or color"], [20, 11, 12, 12, 15])
    skus = [("CND-LAV-8OZ", 24), ("CND-LAV-16OZ", 9), ("CND-CED-8OZ", 18), ("SOP-OAT-4OZ", 31),
            ("SOP-CHR-4OZ", 22), ("SOP-OAT-8OZ", 7), ("EAR-HOOP-GLD", 12), ("EAR-STUD-SLV", 15),
            ("CRD-BDAY-A6", 40), ("CRD-THNX-A6", 28)]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(skus[r - 2]) if r - 2 < len(skus) else [])
        ws[f"C{r}"] = f'=IF(A{r}="","",IFERROR(LEFT(A{r},FIND("-",A{r})-1),"Check SKU"))'
        ws[f"D{r}"] = (f'=IF(A{r}="","",IFERROR(MID(A{r},FIND("-",A{r})+1,'
                       f'FIND("-",A{r},FIND("-",A{r})+1)-FIND("-",A{r})-1),"Check SKU"))')
        ws[f"E{r}"] = f'=IF(A{r}="","",IFERROR(MID(A{r},FIND("-",A{r},FIND("-",A{r})+1)+1,LEN(A{r})),"Check SKU"))'
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["H"].width = 12
    ws["G1"], ws["H1"] = "Category code", "Units sold"
    for cell in (ws["G1"], ws["H1"]):
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    for r, code in enumerate(["CND", "SOP", "EAR", "CRD"], start=2):
        ws[f"G{r}"] = code
        ws[f"G{r}"].fill = INPUT_FILL
        ws[f"H{r}"] = f"=SUMIF($C$2:$C${last},G{r},$B$2:$B${last})"
        ws[f"H{r}"].fill = FORMULA_FILL
    ws["G6"], ws["H6"] = "Total of the categories", "=SUM(H2:H5)"
    ws["G7"], ws["H7"] = "All units on the sheet", f"=SUM(B2:B{last})"
    ws["G8"], ws["H8"] = "Check", '=IF(H6=H7,"OK","CHECK")'
    ws["G6"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Split a Product Code Into Parts With LEFT, MID and FIND",
        "",
        "1. SKUs tab: type or paste codes in column A. Each code has three parts joined by two hyphens, such as CND-LAV-8OZ (category, style, size or color).",
        "2. C2: =IFERROR(LEFT(A2,FIND(\"-\",A2)-1),\"Check SKU\") takes everything before the first hyphen. FIND returns the position of the hyphen.",
        "3. D2 takes the text between the first and second hyphen: =MID(A2,FIND(\"-\",A2)+1,FIND(\"-\",A2,FIND(\"-\",A2)+1)-FIND(\"-\",A2)-1). The second FIND starts looking just after the first hyphen.",
        "4. E2: =MID(A2,FIND(\"-\",A2,FIND(\"-\",A2)+1)+1,LEN(A2)) takes everything after the second hyphen.",
        "5. A code with fewer than two hyphens shows Check SKU. A code with a third hyphen keeps the extra part in column E (CND-LAV-8OZ-GIFT gives 8OZ-GIFT).",
        "6. H2: =SUMIF($C$2:$C$30,G2,$B$2:$B$30) adds units sold for each category code. H8 compares the category total with all units: CHECK means a code did not land in any listed category.",
        "7. A space before or after a code breaks the match. Clean pasted codes with TRIM first.",
        "8. The formulas are ready to row 30; fill them further down if you have more codes.",
        "",
        "Codes and quantities are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-split-sku-code.xlsx")


def last_number_in_column():
    wb = Workbook()
    ws = wb.active
    ws.title = "Log"
    header(ws, ["Date", "Balance", "Note"], [13, 14, 30])
    rows = {
        2: (date(2026, 9, 1), 1000.00, "Opening cash"),
        3: (date(2026, 9, 3), 1085.50, "Saturday market"),
        5: (date(2026, 9, 8), 1210.25, "Online orders"),
        6: (date(2026, 9, 12), 1148.00, "Supply order"),
        9: (date(2026, 9, 19), 1320.40, "Saturday market"),
        10: (date(2026, 9, 24), 1257.90, "Booth fee"),
        12: (date(2026, 10, 2), 1402.75, "Online orders"),
    }
    for r in range(2, 31):
        if r in rows:
            d, b, n = rows[r]
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = d, b, n
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        ws[f"A{r}"].number_format = US_DATE
        ws[f"B{r}"].number_format = USD

    s = wb.create_sheet("Summary")
    s.column_dimensions["A"].width = 40
    s.column_dimensions["B"].width = 20
    s["A1"], s["B1"] = "Latest from the Log tab", "Result"
    for c in s[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    items = [
        ("Latest balance", "=LOOKUP(9.99E+307,Log!B2:B200)", USD),
        ("Latest date", "=LOOKUP(9.99E+307,Log!A2:A200)", US_DATE),
        ("Entries in the balance column", "=COUNT(Log!B2:B200)", "0"),
        ("Latest balance, safe for an empty log", '=IFERROR(LOOKUP(9.99E+307,Log!B2:B200),"No entries")', USD),
        ("Latest date, safe for an empty log", '=IFERROR(LOOKUP(9.99E+307,Log!A2:A200),"No entries")', US_DATE),
    ]
    for r, (label, f, fmt) in enumerate(items, start=2):
        s[f"A{r}"], s[f"B{r}"] = label, f
        s[f"B{r}"].number_format = fmt
        s[f"B{r}"].fill = FORMULA_FILL
        s[f"B{r}"].alignment = Alignment(horizontal="right")

    notes_sheet(wb, [
        "Tidy Tabs: Get the Last Number in a Column",
        "",
        "1. Log tab: type each date in column A and the cash balance in column B. Blank rows between entries are fine. Column C is for notes.",
        "2. Summary!B2: =LOOKUP(9.99E+307,Log!B2:B200) returns the last number in B2:B200. 9.99E+307 is bigger than any balance you will type, so LOOKUP settles on the last number in the range.",
        "3. Summary!B3 does the same for the dates in column A (dates are numbers in a spreadsheet).",
        "4. Summary!B4: =COUNT(Log!B2:B200) counts the numbers in the column. It ignores blank cells and text.",
        "5. Summary!B5 and B6 wrap the formula in IFERROR. An empty column makes LOOKUP return #N/A, and these show No entries instead.",
        "6. Text is ignored. A note typed in column B, or a number stored as text (it is usually left-aligned), is skipped, so the formula returns the number above it. Retype it as a real number.",
        "7. A 0 as the last balance is still returned. The formula looks at B2:B200 only; use a longer range if your log grows past row 200.",
        "8. This finds numbers and dates only. A text version needs a different formula.",
        "",
        "Dates and balances are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-last-number-in-column.xlsx")


def time_zone_converter():
    wb = Workbook()
    ws = wb.active
    ws.title = "Calls"
    header(ws, ["Client", "Client zone", "Your call time", "Hours difference", "Client time", "Day"],
           [24, 14, 22, 12, 22, 13])
    calls = [
        ("Maple Street Bakery", "Eastern", datetime(2026, 10, 14, 9, 0)),
        ("Harbor Yoga Studio", "Central", datetime(2026, 10, 14, 16, 30)),
        ("Birch Design Co", "Mountain", datetime(2026, 10, 15, 10, 0)),
        ("Lantern Print Shop", "Pacific", datetime(2026, 10, 14, 16, 30)),
        ("Cedar Row Realty", "Pacific", datetime(2026, 10, 15, 1, 0)),
        ("Willow Pet Care", "Central", datetime(2026, 10, 15, 0, 15)),
        ("Granite Fitness", "Mountain", datetime(2026, 10, 16, 14, 45)),
        ("Juniper Books", "Pacific", datetime(2026, 10, 16, 18, 0)),
    ]
    last = 20
    for r in range(2, last + 1):
        ws.append(list(calls[r - 2]) if r - 2 < len(calls) else [])
        ws[f"C{r}"].number_format = US_DATETIME
        ws[f"D{r}"] = (f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},Offsets!$A$2:$B$5,2,FALSE)'
                       f'-VLOOKUP($I$2,Offsets!$A$2:$B$5,2,FALSE),"Check zone"))')
        ws[f"E{r}"] = (f'=IF(OR(C{r}="",NOT(ISNUMBER(D{r}))),"",IF(ISNUMBER(C{r}),C{r}+D{r}/24,"Check time"))')
        ws[f"F{r}"] = (f'=IF(NOT(ISNUMBER(E{r})),"",IF(INT(E{r})>INT(C{r}),"Next day",'
                       f'IF(INT(E{r})<INT(C{r}),"Previous day","Same day")))')
        ws[f"E{r}"].number_format = US_DATETIME
        ws[f"D{r}"].number_format = "0.##;-0.##;0"
        for col in "DEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1="=Offsets!$A$2:$A$5", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"B2:B{last}")
    dv2 = DataValidation(type="list", formula1="=Offsets!$A$2:$A$5", allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add("I2")
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['OR(F2="Next day",F2="Previous day")'], fill=ALERT_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 26
    ws.column_dimensions["I"].width = 14
    ws["H2"], ws["I2"] = "Your zone", "Eastern"
    ws["H2"].font = Font(bold=True)
    ws["I2"].fill = INPUT_FILL
    ws["H4"] = "Calls on a different day"
    ws["I4"] = f'=COUNTIF(F2:F{last},"Next day")+COUNTIF(F2:F{last},"Previous day")'
    ws["H5"] = "Calls listed"
    ws["I5"] = f"=COUNT(C2:C{last})"
    ws["I4"].fill = ws["I5"].fill = FORMULA_FILL

    o = wb.create_sheet("Offsets")
    header(o, ["Zone", "Hours from Eastern"], [16, 20])
    for r, (zone, h) in enumerate([("Eastern", 0), ("Central", -1), ("Mountain", -2), ("Pacific", -3)], start=2):
        o[f"A{r}"], o[f"B{r}"] = zone, h
        o[f"A{r}"].fill = o[f"B{r}"].fill = INPUT_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Convert Time Zones for Client Calls",
        "",
        "1. Calls tab: type the client, pick the client zone from the dropdown in B, and type your call date and time in C (for example 10/14/2026 4:30 PM). Rows 2 to 20 have the formulas ready.",
        "2. I2 is your own zone (yellow). Pick Eastern, Central, Mountain or Pacific. The Offsets tab lists each zone as hours from Eastern (Central -1, Mountain -2, Pacific -3).",
        "3. D2 subtracts your zone's offset from the client zone's offset, so Eastern to Pacific is -3 and Pacific to Eastern is 3.",
        "4. E2: =C2+D2/24 is the client's time. One hour is 1/24 of a day, so adding hours divided by 24 moves the date and the time together.",
        "5. F2 compares the dates: Next day, Previous day or Same day, so a call at 1:00 AM for you is not missed as the evening before for the client.",
        "6. Offsets are fixed hours. Most of the US changes clocks on the same dates, so the differences between these four zones stay the same through the year. Arizona (most of it), Hawaii and other places that do not change clocks are not handled, and neither are zones outside these four. Check the current rules before you confirm a call.",
        "7. Type the call time as a real date and time. A time stored as text shows Check time.",
        "",
        "Clients and times are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-time-zone-converter.xlsx")


def count_weekdays_in_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Months"
    header(ws, ["First day", "Last day", "Saturdays", "Sundays", "Day picked in I3", "Booth cost (Saturdays)"],
           [13, 13, 11, 11, 17, 20])
    last = 25
    for r in range(2, last + 1):
        ws.append([date(2026, r - 1, 1)] if r <= 13 else [])
        ws[f"A{r}"].number_format = US_DATE
        ws[f"B{r}"] = f'=IF(A{r}="","",EOMONTH(A{r},0))'
        ws[f"C{r}"] = f'=IF(A{r}="","",NETWORKDAYS.INTL(A{r},B{r},"1111101"))'
        ws[f"D{r}"] = f'=IF(A{r}="","",NETWORKDAYS.INTL(A{r},B{r},"1111110"))'
        ws[f"E{r}"] = f'=IF(A{r}="","",NETWORKDAYS.INTL(A{r},B{r},$I$4))'
        ws[f"F{r}"] = f'=IF(A{r}="","",C{r}*$I$2)'
        ws[f"B{r}"].number_format = US_DATE
        ws[f"F{r}"].number_format = USD
        ws[f"A{r}"].fill = INPUT_FILL
        for col in "BCDEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 34
    ws.column_dimensions["I"].width = 14
    ws.column_dimensions["K"].width = 8
    ws["H1"] = "Settings and totals"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["I2"] = "Booth fee per market day", 25
    ws["I2"].number_format = USD
    ws["H3"], ws["I3"] = "Day to count in column E", "Fri"
    ws["H4"] = "Weekend string for that day"
    ws["I4"] = '=REPT("1",MATCH(I3,$K$2:$K$8,0)-1)&"0"&REPT("1",7-MATCH(I3,$K$2:$K$8,0))'
    ws["I4"].number_format = "@"
    for a in ("I2", "I3"):
        ws[a].fill = INPUT_FILL
    ws["I4"].fill = FORMULA_FILL
    ws["K1"] = "Days"
    ws["K1"].font = Font(bold=True)
    for r, d in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], start=2):
        ws[f"K{r}"] = d
    dv = DataValidation(type="list", formula1="=$K$2:$K$8", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("I3")
    ws["H6"], ws["I6"] = "Saturdays in the year", f"=SUM(C2:C{last})"
    ws["H7"], ws["I7"] = "Sundays in the year", f"=SUM(D2:D{last})"
    ws["H8"], ws["I8"] = "Booth cost for the year", f"=SUM(F2:F{last})"
    ws["I8"].number_format = USD
    ws["H9"] = "Check: Saturdays match one count over the whole span"
    ws["I9"] = (f'=IF(COUNT(A2:A{last})=0,"",IF(I6=NETWORKDAYS.INTL(MIN(A2:A{last}),'
                f'EOMONTH(MAX(A2:A{last}),0),"1111101"),"OK","CHECK"))')
    for a in ("I6", "I7", "I8", "I9"):
        ws[a].fill = FORMULA_FILL
    ws["H9"].alignment = Alignment(wrap_text=True, vertical="top")

    notes_sheet(wb, [
        "Tidy Tabs: Count Market Saturdays in a Month With NETWORKDAYS.INTL",
        "",
        "1. Months tab: column A holds the first day of each month (yellow = input). Rows 2 to 13 are January to December 2026. Rows 14 to 25 are ready for more months.",
        "2. B2: =EOMONTH(A2,0) gives the last day of that month.",
        "3. C2: =NETWORKDAYS.INTL(A2,B2,\"1111101\") counts Saturdays. The string has seven digits for Monday to Sunday. 1 means skip that day, 0 means count it. Only the sixth digit (Saturday) is 0.",
        "4. D2 uses \"1111110\" to count Sundays instead.",
        "5. I3 picks any weekday from a dropdown. I4 builds its string with REPT, and E2 uses it: =NETWORKDAYS.INTL(A2,B2,$I$4).",
        "6. F2: =C2*$I$2 multiplies the Saturdays by the booth fee per market day in I2. It counts calendar Saturdays only. Your real market days, holidays and fees are set by the market.",
        "7. The first date counts too. If the start date is after the end date, the result is negative. The dates must be real dates, not text.",
        "8. I9 checks that the monthly Saturdays add up to one count over the whole span.",
        "",
        "Dates and fees are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-count-weekdays-in-month.xlsx")


def days_to_get_paid():
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Invoice", "Client", "Invoice date", "Paid date", "Days to pay", "Status"],
           [12, 24, 13, 13, 12, 12])
    rows = [
        ("INV-2001", "Maple Street Bakery", date(2026, 7, 1), date(2026, 7, 21)),
        ("INV-2002", "Harbor Yoga Studio", date(2026, 7, 8), date(2026, 8, 12)),
        ("INV-2003", "Birch Design Co", date(2026, 7, 15), date(2026, 7, 29)),
        ("INV-2004", "Lantern Print Shop", date(2026, 7, 22), date(2026, 8, 26)),
        ("INV-2005", "Cedar Row Realty", date(2026, 8, 3), date(2026, 8, 10)),
        ("INV-2006", "Willow Pet Care", date(2026, 8, 10), date(2026, 9, 24)),
        ("INV-2007", "Granite Fitness", date(2026, 8, 19), date(2026, 9, 8)),
        ("INV-2008", "Juniper Books", date(2026, 9, 2), date(2026, 9, 30)),
        ("INV-2009", "Maple Street Bakery", date(2026, 9, 16), None),
        ("INV-2010", "Harbor Yoga Studio", date(2026, 10, 1), None),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"E{r}"] = f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})'
        ws[f"F{r}"] = f'=IF(C{r}="","",IF(D{r}="","Unpaid","Paid"))'
        for col in "CD":
            ws[f"{col}{r}"].number_format = US_DATE
        ws[f"E{r}"].number_format = "0"
        for col in "ABCD":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "EF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"E2:E{last}", FormulaRule(formula=['AND(ISNUMBER(E2),E2>$J$2)'], fill=ALERT_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 30
    ws.column_dimensions["J"].width = 12
    ws["H1"] = "Settings and results"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["J2"] = "Your payment terms in days", 30
    ws["J2"].fill = INPUT_FILL
    results = [
        ("Average days to get paid", f'=IFERROR(AVERAGE(E2:E{last}),"")', "0.0"),
        ("Fastest payment (days)", f'=IF(COUNT(E2:E{last})=0,"",MIN(E2:E{last}))', "0"),
        ("Slowest payment (days)", f'=IF(COUNT(E2:E{last})=0,"",MAX(E2:E{last}))', "0"),
        ("Invoices paid", f"=COUNT(E2:E{last})", "0"),
        ("Invoices still unpaid", f'=COUNTIF(F2:F{last},"Unpaid")', "0"),
        ("Paid later than terms", f'=COUNTIF(E2:E{last},">"&J2)', "0"),
        ("Check: paid + unpaid = invoices", f'=IF(J7+J8=COUNTA(A2:A{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(results, start=4):
        ws[f"H{i}"], ws[f"J{i}"] = label, f
        ws[f"J{i}"].number_format = fmt
        ws[f"J{i}"].fill = FORMULA_FILL
        ws[f"J{i}"].alignment = Alignment(horizontal="right")

    notes_sheet(wb, [
        "Tidy Tabs: Calculate Average Days to Get Paid",
        "",
        "1. Invoices tab: type the invoice number, client, invoice date (A to C) and, once the client pays, the paid date in D. Leave D blank while an invoice is unpaid. Yellow cells are inputs; rows 2 to 30 have formulas ready.",
        "2. E2: =D2-C2 is the days between the invoice date and the paid date. Dates are numbers, so subtracting gives days. A blank paid date leaves E blank.",
        "3. J4: =AVERAGE(E2:E30) is the average days to get paid. AVERAGE skips blank cells, so unpaid invoices do not pull it down to zero.",
        "4. J5 and J6 give the fastest and slowest payment with MIN and MAX. J7 and J8 count paid and unpaid invoices.",
        "5. J2 is your payment terms in days (30 for Net 30). E turns red when an invoice took longer, and J9 counts them.",
        "6. J10 checks that paid plus unpaid equals the number of invoices listed.",
        "7. Enter real dates. A date stored as text (usually left-aligned) can give a wrong number or an error. A paid date earlier than the invoice date gives a negative number; fix the typo.",
        "8. This looks back at invoices that were paid. It does not age open invoices; the receivables aging template does that.",
        "",
        "Clients, dates and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-days-to-get-paid.xlsx")


def early_payment_discount():
    wb = Workbook()
    ws = wb.active
    ws.title = "Discounts"
    header(ws, ["Invoice", "Client", "Invoice date", "Amount", "Paid date", "Discount deadline",
                "Discount amount", "Pay by deadline", "Full amount due date", "Status"],
           [11, 22, 12, 12, 12, 13, 12, 13, 13, 16])
    rows = [
        ("INV-3001", "Maple Street Bakery", date(2026, 9, 1), 1250.00, date(2026, 9, 9)),
        ("INV-3002", "Harbor Yoga Studio", date(2026, 9, 3), 480.00, date(2026, 9, 20)),
        ("INV-3003", "Birch Design Co", date(2026, 9, 8), 2200.00, date(2026, 9, 18)),
        ("INV-3004", "Lantern Print Shop", date(2026, 9, 15), 835.50, date(2026, 9, 30)),
        ("INV-3005", "Cedar Row Realty", date(2026, 9, 22), 960.00, date(2026, 9, 30)),
        ("INV-3006", "Willow Pet Care", date(2026, 10, 1), 312.50, None),
        ("INV-3007", "Granite Fitness", date(2026, 10, 5), 1480.00, None),
        ("INV-3008", "Juniper Books", date(2026, 10, 9), 725.40, None),
    ]
    last = 25
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"F{r}"] = f'=IF(C{r}="","",C{r}+$M$3)'
        ws[f"G{r}"] = f'=IF(OR(C{r}="",NOT(ISNUMBER(D{r}))),"",ROUND(D{r}*$M$2,2))'
        ws[f"H{r}"] = f'=IF(G{r}="","",D{r}-G{r})'
        ws[f"I{r}"] = f'=IF(C{r}="","",C{r}+$M$4)'
        ws[f"J{r}"] = f'=IF(OR(C{r}="",G{r}=""),"",IF(E{r}="","Open",IF(E{r}<=F{r},"Discount earned","Full amount")))'
        for col in "CEFI":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "DGH":
            ws[f"{col}{r}"].number_format = USD
        for col in "ABCDE":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "FGHIJ":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"J2:J{last}", FormulaRule(formula=['J2="Full amount"'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"J2:J{last}", FormulaRule(formula=['J2="Discount earned"'], fill=DONE_FILL))
    ws.column_dimensions["K"].width = 4
    ws.column_dimensions["L"].width = 32
    ws.column_dimensions["M"].width = 14
    ws["L1"] = "Terms and totals"
    ws["L1"].font = Font(bold=True)
    ws["L2"], ws["M2"] = "Discount rate", 0.02
    ws["L3"], ws["M3"] = "Discount days", 10
    ws["L4"], ws["M4"] = "Net days (full amount due)", 30
    ws["M2"].number_format = "0.0%"
    for a in ("M2", "M3", "M4"):
        ws[a].fill = INPUT_FILL
    totals = [
        ("Discounts taken by clients", f'=SUMIF(J2:J{last},"Discount earned",G2:G{last})', USD),
        ("Discounts missed", f'=SUMIF(J2:J{last},"Full amount",G2:G{last})', USD),
        ("Discount still available (open)", f'=SUMIF(J2:J{last},"Open",G2:G{last})', USD),
        ("Open invoices", f'=COUNTIF(J2:J{last},"Open")', "0"),
    ]
    for i, (label, f, fmt) in enumerate(totals, start=6):
        ws[f"L{i}"], ws[f"M{i}"] = label, f
        ws[f"M{i}"].number_format = fmt
        ws[f"M{i}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Calculate an Early Payment Discount (2/10 Net 30)",
        "",
        "1. Discounts tab: type the invoice, client, invoice date and amount (A to D). When a client pays, type the paid date in E; leave it blank while open. Rows 2 to 25 have formulas ready.",
        "2. M2, M3 and M4 hold your terms: discount rate (2%), discount days (10) and net days (30). 2/10 net 30 means 2% off if paid within 10 days, otherwise the full amount in 30. Change them for terms like 1/15 net 45.",
        "3. F2: =C2+$M$3 is the last day to pay and still take the discount.",
        "4. G2: =ROUND(D2*$M$2,2) is the discount in dollars. H2: =D2-G2 is what the client pays by the deadline. I2: =C2+$M$4 is the date the full amount is due.",
        "5. J2 compares the paid date with the deadline. A payment on the deadline day still earns the discount. Blank paid date shows Open.",
        "6. M6 to M9 total the discounts taken, missed and still available.",
        "7. Dates must be real dates, not text. The discount is calendar days, not business days.",
        "8. This is arithmetic from the terms you type. The terms that count are the ones in your contract or invoice; the template does not decide what a client owes.",
        "",
        "Clients, dates and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-early-payment-discount.xlsx")


def sell_through_rate():
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    header(ws, ["Product", "Units received", "Units sold", "Units left", "Sell through", "Status"],
           [30, 14, 12, 12, 13, 14])
    rows = [
        ("Lavender soy candle, 8 oz", 48, 41),
        ("Cedar soy candle, 8 oz", 48, 30),
        ("Oatmeal goat milk soap", 120, 96),
        ("Charcoal soap bar", 60, 21),
        ("Brass hoop earrings", 30, 12),
        ("Letterpress birthday card", 100, 78),
        ("Linen tote bag", 25, 25),
        ("Gift tag set", 80, 18),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"D{r}"] = f'=IF(OR(B{r}="",C{r}=""),"",B{r}-C{r})'
        ws[f"E{r}"] = f'=IF(OR(NOT(ISNUMBER(B{r})),NOT(ISNUMBER(C{r}))),"",IF(B{r}<=0,"",C{r}/B{r}))'
        ws[f"F{r}"] = (f'=IF(E{r}="","",IF(E{r}>1,"Check counts",IF(E{r}>=$I$2,"Strong",'
                       f'IF(E{r}<$I$3,"Slow","OK"))))')
        ws[f"E{r}"].number_format = "0.0%"
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['F2="Slow"'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['F2="Strong"'], fill=DONE_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 30
    ws.column_dimensions["I"].width = 12
    ws["H1"] = "Settings and totals"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["I2"] = "Strong at or above", 0.7
    ws["H3"], ws["I3"] = "Slow below", 0.4
    for a in ("I2", "I3"):
        ws[a].number_format = "0%"
        ws[a].fill = INPUT_FILL
    totals = [
        ("Units received", f"=SUM(B2:B{last})", "0"),
        ("Units sold", f"=SUM(C2:C{last})", "0"),
        ("Overall sell through", f'=IF(I5=0,"",I6/I5)', "0.0%"),
        ("Products marked Slow", f'=COUNTIF(F2:F{last},"Slow")', "0"),
        ("Units left on those", f'=SUMIF(F2:F{last},"Slow",D2:D{last})', "0"),
    ]
    for i, (label, f, fmt) in enumerate(totals, start=5):
        ws[f"H{i}"], ws[f"I{i}"] = label, f
        ws[f"I{i}"].number_format = fmt
        ws[f"I{i}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Calculate Sell Through Rate",
        "",
        "1. Products tab: type the product, the units you received or made (B) and the units sold so far (C). Rows 2 to 30 have formulas ready.",
        "2. D2: =B2-C2 is the units left.",
        "3. E2: =C2/B2 is the sell through rate: units sold divided by units received. 41 of 48 candles sold is 85.4%. Format the cell as a percent.",
        "4. E2 stays blank when units received is 0 or blank, so there is no divide by zero error.",
        "5. I2 and I3 are your own cutoffs (yellow). F2 shows Strong at or above I2, Slow below I3, otherwise OK. The 70% and 40% are examples, not industry rules.",
        "6. If units sold is more than units received, F2 shows Check counts. Fix a typo or add the missing stock.",
        "7. I7 is the overall rate: total sold divided by total received, not the average of the row percents.",
        "8. Compare products over the same time period. This sheet counts units, not dollars.",
        "",
        "Products and counts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-sell-through-rate.xlsx")


def boxes_needed():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Item", "Units to ship", "Units per box", "Boxes needed", "Units in last box",
                "Empty slots", "Box cost"],
           [28, 12, 12, 12, 13, 11, 12])
    rows = [
        ("Soy candle 8 oz", 50, 12),
        ("Soap bars", 30, 8),
        ("Earring cards", 17, 20),
        ("Ceramic mugs", 26, 6),
        ("Linen tote bags", 48, 12),
        ("Birthday cards", 130, 25),
        ("Gift tag sets", 0, 40),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ok = f'AND(ISNUMBER(B{r}),ISNUMBER(C{r}),C{r}>0)'
        ws[f"D{r}"] = f'=IF({ok},ROUNDUP(B{r}/C{r},0),"")'
        ws[f"E{r}"] = f'=IF(OR(D{r}="",B{r}=0),"",B{r}-(D{r}-1)*C{r})'
        ws[f"F{r}"] = f'=IF(D{r}="","",D{r}*C{r}-B{r})'
        ws[f"G{r}"] = f'=IF(D{r}="","",D{r}*$J$2)'
        ws[f"G{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DEFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.column_dimensions["H"].width = 4
    ws.column_dimensions["I"].width = 28
    ws.column_dimensions["J"].width = 12
    ws["I1"] = "Cost and totals"
    ws["I1"].font = Font(bold=True)
    ws["I2"], ws["J2"] = "Cost per box", 1.45
    ws["J2"].number_format = USD
    ws["J2"].fill = INPUT_FILL
    totals = [
        ("Total boxes", f"=SUM(D2:D{last})", "0"),
        ("Total box cost", f"=SUM(G2:G{last})", USD),
        ("Total empty slots", f"=SUM(F2:F{last})", "0"),
    ]
    for i, (label, f, fmt) in enumerate(totals, start=4):
        ws[f"I{i}"], ws[f"J{i}"] = label, f
        ws[f"J{i}"].number_format = fmt
        ws[f"J{i}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: Calculate How Many Boxes You Need",
        "",
        "1. Orders tab: type the item, the units to ship (B) and how many fit in one box (C). Rows 2 to 30 have formulas ready.",
        "2. D2: =ROUNDUP(B2/C2,0) is the boxes needed. ROUNDUP always rounds up, so 50 candles at 12 per box (4.17) needs 5 boxes. A plain ROUND would give 4 and leave candles without a box.",
        "3. E2: =B2-(D2-1)*C2 is how many units go in the last box. When the units divide evenly the last box is simply full.",
        "4. F2: =D2*C2-B2 is the empty slots across all boxes.",
        "5. G2: =D2*$J$2 is the box cost. J2 is your price per box (yellow).",
        "6. If units per box is 0, blank or text, D to G stay blank. If units to ship is 0, D shows 0 boxes.",
        "7. Same size boxes only. It does not check weight, box dimensions or carrier rules, and one box size per row.",
        "",
        "Items, box sizes and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-boxes-needed.xlsx")


def median_order_value():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order", "Date", "Order total"], [12, 13, 14])
    totals_list = [24.50, 38.00, 42.75, 18.00, 55.25, 31.50, 640.00, 27.00, 46.50, 35.25, 29.00, 62.00, 33.75]
    last = 100
    for r in range(2, last + 1):
        i = r - 2
        if i < len(totals_list):
            ws.append([f"ORD-{4001 + i}", date(2026, 9, 1 + 2 * i), totals_list[i]])
        else:
            ws.append([])
        ws[f"B{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 34
    ws.column_dimensions["F"].width = 14
    ws["E1"] = "Settings and results"
    ws["E1"].font = Font(bold=True)
    ws["E2"], ws["F2"] = "Ignore orders above (used in F9, F10)", 500
    ws["F2"].number_format = USD
    ws["F2"].fill = INPUT_FILL
    res = [
        ("Orders counted", f"=COUNT(C2:C{last})", "0"),
        ("Average order value", f'=IFERROR(AVERAGE(C2:C{last}),"")', USD),
        ("Median order value", f'=IFERROR(MEDIAN(C2:C{last}),"")', USD),
        ("Average minus median", f'=IF(F5="","",F5-F6)', USD),
        ("Largest order", f'=IF(F4=0,"",MAX(C2:C{last}))', USD),
        ("Average without orders above F2", f'=IFERROR(AVERAGEIF(C2:C{last},"<="&F2),"")', USD),
        ("Orders above F2", f'=COUNTIF(C2:C{last},">"&F2)', "0"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=4):
        ws[f"E{i}"], ws[f"F{i}"] = label, f
        ws[f"F{i}"].number_format = fmt
        ws[f"F{i}"].fill = FORMULA_FILL
        ws[f"F{i}"].alignment = Alignment(horizontal="right")

    notes_sheet(wb, [
        "Tidy Tabs: Find the Median Order Value",
        "",
        "1. Orders tab: type the order number, date and order total in A to C (yellow). Totals can go down to row 100.",
        "2. F5: =AVERAGE(C2:C100) adds every order and divides by the count. One big wholesale order pulls it up.",
        "3. F6: =MEDIAN(C2:C100) is the middle order when the totals are sorted. With an even count it is the average of the two middle orders. A single large order barely moves it.",
        "4. F7 shows how far the average sits above the median. A large gap means a few big orders are skewing the average.",
        "5. F9 and F10 are optional: type a cutoff in F2 and F9 averages only orders at or below it, F10 counts the orders above it. The cutoff is your call; it is not a rule.",
        "6. Blank cells and text are ignored by AVERAGE and MEDIAN. A total typed as text (usually left-aligned) is skipped, so retype it as a number. With no orders, F5 and F6 stay blank.",
        "7. Order totals only. It does not remove refunds or tax; enter the totals you want to measure.",
        "",
        "Orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-median-order-value.xlsx")

def week_ending_date():
    wb = Workbook()
    ws = wb.active
    ws.title = "Log"
    header(ws, ["Date", "Client", "Hours", "Week ending"], [13, 24, 10, 14])
    rows = [
        (date(2026, 9, 28), "Maple Street Bakery", 3.5),
        (date(2026, 9, 29), "Harbor Yoga Studio", 4.0),
        (date(2026, 9, 30), "Maple Street Bakery", 2.5),
        (date(2026, 10, 1), "Birch Design Co", 6.0),
        (date(2026, 10, 2), "Birch Design Co", 5.5),
        (date(2026, 10, 3), "Harbor Yoga Studio", 2.0),
        (date(2026, 10, 4), "Lantern Print Shop", 1.5),
        (date(2026, 10, 5), "Lantern Print Shop", 7.0),
        (date(2026, 10, 7), "Maple Street Bakery", 4.25),
        (date(2026, 10, 9), "Birch Design Co", 3.0),
        (date(2026, 10, 12), "Harbor Yoga Studio", 5.0),
        (date(2026, 10, 14), "Lantern Print Shop", 2.75),
    ]
    last = 60
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"D{r}"] = f'=IF(ISNUMBER(A{r}),A{r}+MOD($G$2-WEEKDAY(A{r},1),7),"")'
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = "0.00"
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        ws[f"D{r}"].fill = FORMULA_FILL
    ws.column_dimensions["E"].width = 4
    ws.column_dimensions["F"].width = 34
    ws.column_dimensions["G"].width = 12
    ws["F1"] = "Setting"
    ws["F1"].font = Font(bold=True)
    ws["F2"], ws["G2"] = "Week ends on (1 Sun, 2 Mon ... 6 Fri, 7 Sat)", 6
    ws["G2"].fill = INPUT_FILL
    ws["F3"] = "Week ends on"
    ws["G3"] = '=IF(AND(ISNUMBER(G2),G2>=1,G2<=7),CHOOSE(G2,"Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"),"Enter 1 to 7")'
    ws["G3"].fill = FORMULA_FILL
    ws["F4"], ws["G4"] = "Total hours in log", f"=SUM(C2:C{last})"
    ws["G4"].number_format = "0.00"
    ws["G4"].fill = FORMULA_FILL
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="7", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("G2")

    wk = wb.create_sheet("Weeks")
    header(wk, ["Week ending", "Hours"], [14, 10])
    wlast = 27
    for r in range(2, wlast + 1):
        if r == 2:
            wk["A2"] = f'=IF(COUNT(Log!D2:D{last})=0,"",MIN(Log!D2:D{last}))'
        else:
            wk[f"A{r}"] = f'=IF(A{r-1}="","",IF(A{r-1}+7>MAX(Log!D2:D{last}),"",A{r-1}+7))'
        wk[f"B{r}"] = f'=IF(A{r}="","",SUMIFS(Log!C2:C{last},Log!D2:D{last},A{r}))'
        wk[f"A{r}"].number_format = US_DATE
        wk[f"B{r}"].number_format = "0.00"
        wk[f"A{r}"].fill = FORMULA_FILL
        wk[f"B{r}"].fill = FORMULA_FILL
    wk.column_dimensions["C"].width = 4
    wk.column_dimensions["D"].width = 34
    wk.column_dimensions["E"].width = 12
    wk["D1"] = "Check"
    wk["D1"].font = Font(bold=True)
    wk["D2"], wk["E2"] = "Hours on this tab", f"=SUM(B2:B{wlast})"
    wk["D3"], wk["E3"] = "Hours in the log", f"=SUM(Log!C2:C{last})"
    wk["D4"], wk["E4"] = "Check: they match", '=IF(ROUND(E2-E3,6)=0,"OK","CHECK")'
    for a in ("E2", "E3"):
        wk[a].number_format = "0.00"
    for a in ("E2", "E3", "E4"):
        wk[a].fill = FORMULA_FILL
        wk[a].alignment = Alignment(horizontal="right")
    wk.conditional_formatting.add("E4", FormulaRule(formula=['E4="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Find the Week Ending Date for Any Date",
        "",
        "1. Log tab: type the date, client and hours in A to C (yellow). Rows 2 to 60 have formulas ready.",
        "2. G2 is the weekday that ends your week: 1 Sunday, 2 Monday, 3 Tuesday, 4 Wednesday, 5 Thursday, 6 Friday, 7 Saturday. It starts at 6 (Friday).",
        "3. D2: =A2+MOD($G$2-WEEKDAY(A2,1),7). WEEKDAY(A2,1) gives 1 for Sunday to 7 for Saturday. MOD counts the days forward to the week ending day, from 0 to 6.",
        "4. A date that already falls on the week ending day stays the same. The day after it jumps ahead 6 days.",
        "5. Weeks tab: A2 is the earliest week ending in the log and each row below adds 7 days. B uses SUMIFS to add the hours whose week ending matches.",
        "6. Weeks!E4 checks that the hours on the Weeks tab equal the hours in the log. CHECK means a row was left out, for example a date typed as text, or more weeks than the 26 rows listed.",
        "7. Enter real dates. A date stored as text (usually left-aligned) gets no week ending, and its hours show up as CHECK. A blank date row stays blank.",
        "8. Date math only. It does not apply pay period rules, overtime or time zones.",
        "",
        "Clients, dates and hours are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-week-ending-date.xlsx")


def free_shipping_threshold():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order", "Subtotal", "Delivery", "Status", "Amount to add"], [12, 12, 12, 12, 14])
    rows = [
        ("ORD-5001", 32.50, "Ship"),
        ("ORD-5002", 50.00, "Ship"),
        ("ORD-5003", 49.99, "Ship"),
        ("ORD-5004", 78.25, "Ship"),
        ("ORD-5005", 24.00, "Pickup"),
        ("ORD-5006", 45.00, "Ship"),
        ("ORD-5007", 112.40, "Ship"),
        ("ORD-5008", 18.75, "Ship"),
        ("ORD-5009", 61.00, "Pickup"),
        ("ORD-5010", 41.50, "Ship"),
    ]
    last = 31
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"D{r}"] = (f'=IF(OR(NOT(ISNUMBER(B{r})),C{r}=""),"",IF(C{r}="Pickup","Pickup",'
                       f'IF(AND(C{r}="Ship",B{r}>=$I$2),"Free","Pay")))')
        ws[f"E{r}"] = f'=IF(OR(D{r}="Pay",D{r}="Free"),MAX(0,ROUND($I$2-B{r},2)),"")'
        ws[f"B{r}"].number_format = USD
        ws[f"E{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1='"Ship,Pickup"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C2:C{last}")
    ws.conditional_formatting.add(f"D2:D{last}", FormulaRule(formula=['D2="Free"'], fill=DONE_FILL))
    ws.column_dimensions["F"].width = 4
    ws.column_dimensions["G"].width = 34
    ws.column_dimensions["I"].width = 12
    ws["G1"] = "Settings and results"
    ws["G1"].font = Font(bold=True)
    ws["G2"], ws["I2"] = "Free shipping at or above", 50
    ws["G3"], ws["I3"] = "Close = within this much", 10
    for a in ("I2", "I3"):
        ws[a].number_format = USD
        ws[a].fill = INPUT_FILL
    res = [
        ("Orders that qualify (Free)", f'=COUNTIF(D2:D{last},"Free")', "0"),
        ("Orders that pay shipping", f'=COUNTIF(D2:D{last},"Pay")', "0"),
        ("Pay orders within the close gap", f'=COUNTIFS(D2:D{last},"Pay",E2:E{last},"<="&I3)', "0"),
        ("Pickup orders", f'=COUNTIF(D2:D{last},"Pickup")', "0"),
        ("Check: all rows counted", f'=IF(I4+I5+I7=COUNTA(A2:A{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=4):
        ws[f"G{i}"], ws[f"I{i}"] = label, f
        ws[f"I{i}"].number_format = fmt
        ws[f"I{i}"].fill = FORMULA_FILL
        ws[f"I{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("I8", FormulaRule(formula=['I8="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Check a Free Shipping Threshold",
        "",
        "1. Orders tab: type the order number, subtotal and delivery (Ship or Pickup) in A to C (yellow). Rows 2 to 31 have formulas ready.",
        "2. I2 is your free shipping threshold. I3 is how close counts as close. Both are examples; type your own.",
        "3. D2: =IF(AND(C2=\"Ship\",B2>=$I$2),\"Free\",\"Pay\"), wrapped so Pickup rows show Pickup. A subtotal exactly equal to I2 qualifies (>=); one cent under does not.",
        "4. E2: =MAX(0,$I$2-B2) is how much more the customer must add to reach free shipping. It is 0 when the order already qualifies. ROUND(...,2) keeps cents exact.",
        "5. I4 to I7 count Free, Pay, Pay orders within the close gap (the customers a nudge could convert) and Pickup rows.",
        "6. I8 checks that Free + Pay + Pickup equals the number of order numbers listed. CHECK usually means a subtotal typed as text or a blank delivery cell.",
        "7. A subtotal typed as text (usually left-aligned) or a blank subtotal gets no status.",
        "8. Your own threshold only. It does not know carrier rates or marketplace rules.",
        "",
        "Orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-free-shipping-threshold.xlsx")


def forecast_next_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    header(ws, ["Month number", "Month", "Sales"], [14, 12, 13])
    sales = [1200, 1350, 1500, 1480, 1620, 1710, 1690, 1850]
    last = 25
    for r in range(2, last + 1):
        i = r - 2
        ws.append([i + 1, date(2026, 1 + i, 1), sales[i]] if i < len(sales) else [])
        ws[f"B{r}"].number_format = "mmm yyyy"
        ws[f"C{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
    ws.column_dimensions["D"].width = 4
    for col, w in (("E", 22), ("F", 14), ("G", 14)):
        ws.column_dimensions[col].width = w
    ws["E1"], ws["F1"], ws["G1"] = "Forecast", "Month", "Sales estimate"
    for a in ("E1", "F1", "G1"):
        ws[a].font = Font(bold=True)
    X, Y = f"$A$2:$A${last}", f"$C$2:$C${last}"
    for k in (1, 2, 3):
        r = 1 + k
        ws[f"E{r}"] = f'=IF(COUNT({X})=0,"",MAX({X})+{k})'
        ws[f"F{r}"] = f'=IF(OR(E{r}="",COUNT($B$2:$B${last})=0),"",EDATE(MAX($B$2:$B${last}),E{r}-MAX({X})))'
        ws[f"G{r}"] = f'=IF(OR(E{r}="",COUNT({Y})<2),"",FORECAST(E{r},{Y},{X}))'
        ws[f"F{r}"].number_format = "mmm yyyy"
        ws[f"G{r}"].number_format = USD
        for col in "EFG":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"E{r}"].number_format = "0"
    ws["E6"], ws["G6"] = "Trend per month (SLOPE)", f'=IF(COUNT({Y})<2,"",SLOPE({Y},{X}))'
    ws["E7"], ws["G7"] = "Months of data", f"=COUNT({Y})"
    ws["E8"], ws["G8"] = "Hand calculation, month 1", (
        f'=IF(COUNT({Y})<2,"",(AVERAGE({Y})+(SUMPRODUCT({X},{Y})-COUNT({Y})*AVERAGE({X})*AVERAGE({Y}))'
        f'/(SUMSQ({X})-COUNT({X})*AVERAGE({X})^2)*(E2-AVERAGE({X}))))')
    ws["E9"], ws["G9"] = "Check: FORECAST matches", '=IF(G2="","",IF(ROUND(G2-G8,6)=0,"OK","CHECK"))'
    ws["G6"].number_format = USD
    ws["G8"].number_format = USD
    ws["G7"].number_format = "0"
    for r in range(6, 10):
        ws[f"G{r}"].fill = FORMULA_FILL
        ws[f"G{r}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("G9", FormulaRule(formula=['G9="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Forecast Next Month's Sales With FORECAST",
        "",
        "1. Sales tab: type a month number (1, 2, 3 ...), the month and the sales in A to C (yellow). Rows 2 to 25 are ready for up to 24 months. Keep the month numbers in order, with no gaps.",
        "2. G2: =FORECAST(E2,$C$2:$C$25,$A$2:$A$25). E2 is the next month number (the last one plus 1). FORECAST fits a straight line through your past months and reads the sales value off that line. E3 and E4 give the next two months.",
        "3. G6: =SLOPE($C$2:$C$25,$A$2:$A$25) is how many dollars the line rises or falls each month.",
        "4. G8 repeats the first projection by the least squares formula typed out by hand, and G9 shows OK when FORECAST matches it. CHECK means the month numbers have gaps or text.",
        "5. A straight line is only an estimate drawn from the past. It does not know about seasons, sales or promotions, and it is not a prediction. Use it as one input.",
        "6. Rows with a blank or text month number or sales amount are skipped as a pair by LibreOffice Calc. Retype text as a number.",
        "7. FORECAST needs at least 2 months; with fewer the results stay blank. If the two ranges are different sizes the formula returns an error (#VALUE! in LibreOffice Calc); make both ranges the same length.",
        "8. FORECAST.LINEAR is the newer name in Excel and Google Sheets. This template uses the older FORECAST, which works in both.",
        "",
        "Months and sales are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-forecast-next-month.xlsx")


def client_tenure():
    wb = Workbook()
    ws = wb.active
    ws.title = "Clients"
    header(ws, ["Client", "First project date", "Years", "Months", "Time with you", "Total months"],
           [26, 16, 8, 9, 14, 13])
    rows = [
        ("Maple Street Bakery", date(2021, 3, 15)),
        ("Harbor Yoga Studio", date(2022, 11, 2)),
        ("Birch Design Co", date(2024, 2, 29)),
        ("Lantern Print Shop", date(2025, 8, 20)),
        ("Cedar Row Realty", date(2023, 10, 16)),
        ("Willow Pet Care", date(2019, 6, 30)),
        ("Granite Fitness", date(2026, 1, 5)),
        ("Juniper Books", date(2020, 12, 31)),
        ("Copper Kettle Cafe", date(2026, 10, 1)),
        ("Pine Hill Dental", date(2018, 4, 9)),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"C{r}"] = (f'=IF(B{r}="","",IF(NOT(ISNUMBER(B{r})),"Check date",'
                       f'IF(B{r}>$I$2,"Check date",DATEDIF(B{r},$I$2,"y"))))')
        ws[f"D{r}"] = f'=IF(ISNUMBER(C{r}),DATEDIF(B{r},$I$2,"ym"),"")'
        ws[f"E{r}"] = f'=IF(ISNUMBER(C{r}),C{r}&" yr "&D{r}&" mo","")'
        ws[f"F{r}"] = f'=IF(ISNUMBER(C{r}),C{r}*12+D{r},"")'
        ws[f"B{r}"].number_format = US_DATE
        for col in "AB":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "CDEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"C2:C{last}", FormulaRule(formula=['C2="Check date"'], fill=ALERT_FILL))
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 34
    ws.column_dimensions["I"].width = 12
    ws["H1"] = "Settings and results"
    ws["H1"].font = Font(bold=True)
    ws["H2"], ws["I2"] = "As of date", date(2026, 10, 16)
    ws["H3"], ws["I3"] = "Count clients with at least (years)", 3
    ws["I2"].number_format = US_DATE
    for a in ("I2", "I3"):
        ws[a].fill = INPUT_FILL
    res = [
        ("Clients with a valid date", f"=COUNT(C2:C{last})", "0"),
        ("Clients at or over I3 years", f'=COUNTIF(C2:C{last},">="&I3)', "0"),
        ("Longest time (months)", f'=IF(COUNT(F2:F{last})=0,"",MAX(F2:F{last}))', "0"),
        ("Rows to fix (Check date)", f'=COUNTIF(C2:C{last},"Check date")', "0"),
        ("Check: all clients counted", f'=IF(I4+I7=COUNTA(A2:A{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=4):
        ws[f"H{i}"], ws[f"I{i}"] = label, f
        ws[f"I{i}"].number_format = fmt
        ws[f"I{i}"].fill = FORMULA_FILL
        ws[f"I{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("I8", FormulaRule(formula=['I8="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: How Long Has a Client Been With You",
        "",
        "1. Clients tab: type the client and the first project date in A and B (yellow). Rows 2 to 30 have formulas ready.",
        "2. I2 is the as of date. It is typed, not TODAY(), so the results do not change when you reopen the file. Replace it with =TODAY() if you want them to.",
        "3. C2: =DATEDIF(B2,$I$2,\"y\") counts complete years. D2: =DATEDIF(B2,$I$2,\"ym\") counts the months left over after those years.",
        "4. E2 joins them into text like 5 yr 7 mo, and F2 is the total months (years times 12 plus months).",
        "5. DATEDIF counts complete months only: a start on the 20th is not a full month until the 20th of the next month.",
        "6. DATEDIF is not listed in the Excel function picker, but it works when typed. Google Sheets lists it.",
        "7. DATEDIF returns #NUM! when the start date is after the as of date. This sheet shows Check date instead, and so for a date typed as text. A blank date row stays blank.",
        "8. I3 is a number of years you type; I5 counts the clients at or over it.",
        "",
        "Clients and dates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-client-tenure.xlsx")


def incomplete_rows():
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"
    header(ws, ["Name", "Email", "Phone", "Street", "ZIP", "Blank fields", "Status", "First missing"],
           [22, 28, 15, 24, 8, 9, 11, 14])
    rows = [
        ("Ava Thompson", "ava.t@example.com", "617-555-0101", "12 Elm St", "02108"),
        ("Ben Ortiz", "", "201-555-0142", "88 Pine Ave", "07030"),
        ("Carla Nguyen", "carla.n@example.com", "", "5 Harbor Rd", "02210"),
        ("Dev Patel", "dev.p@example.com", "973-555-0188", "", "07102"),
        ("Erin Walsh", "erin.w@example.com", "617-555-0133", "41 Cedar Ln", ""),
        ("Frank Lee", "", "", "9 Mill St", "01002"),
        ("Gina Rossi", "gina.r@example.com", "201-555-0177", "230 Oak Dr", "07601"),
        ("", "hal.k@example.com", "973-555-0120", "17 Birch Ct", "07042"),
        ("Iris Chen", "iris.c@example.com", "617-555-0165", "63 Maple St", "02139"),
        ("Jon Baker", "", "", "", ""),
        ("Kim Alvarez", "kim.a@example.com", "201-555-0111", "70 River Rd", "07024"),
        ("Liam Burke", "liam.b@example.com", "", "3 Lake Ave", "01851"),
    ]
    last = 30
    for r in range(2, last + 1):
        ws.append([v if v != "" else None for v in rows[r - 2]] if r - 2 < len(rows) else [])
        ws[f"F{r}"] = f'=IF(COUNTA(A{r}:E{r})=0,"",COUNTBLANK(A{r}:E{r}))'
        ws[f"G{r}"] = f'=IF(F{r}="","",IF(F{r}=0,"Complete","Missing"))'
        ws[f"H{r}"] = (f'=IF(OR(F{r}="",F{r}=0),"",IF(A{r}="","Name",IF(B{r}="","Email",'
                       f'IF(C{r}="","Phone",IF(D{r}="","Street","ZIP")))))')
        ws[f"E{r}"].number_format = "@"
        for col in "ABCDE":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "FGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=['G2="Missing"'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=['G2="Complete"'], fill=DONE_FILL))
    ws.column_dimensions["I"].width = 4
    ws.column_dimensions["J"].width = 30
    ws.column_dimensions["K"].width = 12
    ws["J1"] = "Totals"
    ws["J1"].font = Font(bold=True)
    res = [
        ("Rows with data", f"=COUNT(F2:F{last})", "0"),
        ("Complete rows", f'=COUNTIF(G2:G{last},"Complete")', "0"),
        ("Rows missing something", f'=COUNTIF(G2:G{last},"Missing")', "0"),
        ("Rows missing a ZIP", f'=SUMPRODUCT((F2:F{last}<>"")*(E2:E{last}=""))', "0"),
        ("Check: blanks add up", f'=IF(SUM(F2:F{last})=COUNT(F2:F{last})*5-COUNTA(A2:E{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=2):
        ws[f"J{i}"], ws[f"K{i}"] = label, f
        ws[f"K{i}"].number_format = fmt
        ws[f"K{i}"].fill = FORMULA_FILL
        ws[f"K{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("K6", FormulaRule(formula=['K6="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Find Rows With Missing Information",
        "",
        "1. Customers tab: paste or type name, email, phone, street and ZIP in A to E (yellow). Rows 2 to 30 have formulas ready. The ZIP column is formatted as text so 02108 keeps its leading zero.",
        "2. F2: =COUNTBLANK(A2:E2) counts the empty cells in the row. A row with no data at all stays blank here, so unused rows do not show 5.",
        "3. G2 shows Complete when F2 is 0 and Missing otherwise. H2 names the first empty field, checked in the order Name, Email, Phone, Street, ZIP.",
        "4. K3 to K5 count complete rows, rows missing something and rows with no ZIP. K6 checks that the blanks counted per row add up to the cells actually filled.",
        "5. COUNTBLANK counts cells that are really empty. A cell holding only a space is not empty, so it counts as filled. Use the TRIM and clean list template to remove stray spaces first.",
        "6. A cell with a formula that returns \"\" counts as blank in LibreOffice Calc and Excel; K6 shows CHECK in that case because COUNTA treats the cell as filled.",
        "7. This counts empty cells only. It does not check that an email or address is correct.",
        "",
        "Names, emails, phone numbers and addresses are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-incomplete-rows.xlsx")


def overlapping_bookings():
    wb = Workbook()
    ws = wb.active
    ws.title = "Bookings"
    header(ws, ["Booking ID", "Space", "Client", "Start date", "End date", "Conflicts", "Status", "Cross-check"],
           [12, 11, 24, 12, 12, 10, 11, 11])
    rows = [
        ("BK-101", "Booth A", "Maple Street Bakery", date(2026, 10, 3), date(2026, 10, 5)),
        ("BK-102", "Booth A", "Harbor Yoga Studio", date(2026, 10, 5), date(2026, 10, 7)),
        ("BK-103", "Booth A", "Birch Design Co", date(2026, 10, 12), date(2026, 10, 20)),
        ("BK-104", "Booth A", "Lantern Print Shop", date(2026, 10, 14), date(2026, 10, 16)),
        ("BK-105", "Booth B", "Cedar Row Realty", date(2026, 10, 3), date(2026, 10, 5)),
        ("BK-106", "Booth B", "Willow Pet Care", date(2026, 10, 10), date(2026, 10, 14)),
        ("BK-107", "Booth B", "Granite Fitness", date(2026, 10, 13), date(2026, 10, 17)),
        ("BK-108", "Booth B", "Juniper Books", date(2026, 10, 16), date(2026, 10, 19)),
        ("BK-109", "Studio", "Copper Kettle Cafe", date(2026, 10, 8), date(2026, 10, 9)),
        ("BK-110", "Studio", "Pine Hill Dental", date(2026, 10, 8), date(2026, 10, 9)),
        ("BK-111", "Studio", "Maple Street Bakery", date(2026, 10, 21), date(2026, 10, 23)),
        ("BK-112", "Studio", "Harbor Yoga Studio", date(2026, 10, 24), date(2026, 10, 24)),
    ]
    last = 40
    sp, st, en = f"$B$2:$B${last}", f"$D$2:$D${last}", f"$E$2:$E${last}"
    for r in range(2, last + 1):
        ws.append(list(rows[r - 2]) if r - 2 < len(rows) else [])
        ws[f"F{r}"] = (
            f'=IF(A{r}="","",IF(OR(B{r}="",NOT(ISNUMBER(D{r})),NOT(ISNUMBER(E{r}))),"",IF(E{r}<D{r},"",'
            f'IF($K$2="Yes",COUNTIFS({sp},B{r},{st},"<="&E{r},{en},">="&D{r})-1,'
            f'COUNTIFS({sp},B{r},{st},"<"&E{r},{en},">"&D{r})-IF(E{r}>D{r},1,0)))))')
        ws[f"G{r}"] = f'=IF(A{r}="","",IF(F{r}="","Check",IF(F{r}>0,"Conflict","OK")))'
        ws[f"H{r}"] = (
            f'=IF(F{r}="","",IF($K$2="Yes",SUMPRODUCT(({sp}=B{r})*({st}<=E{r})*({en}>=D{r}))-1,'
            f'SUMPRODUCT(({sp}=B{r})*({st}<E{r})*({en}>D{r}))-IF(E{r}>D{r},1,0)))')
        for col in "DE":
            ws[f"{col}{r}"].number_format = US_DATE
        for col in "ABCDE":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "FGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1='"Booth A,Booth B,Studio"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"B2:B{last}")
    dv2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add("K2")
    ws.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=['OR(G2="Conflict",G2="Check")'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"G2:G{last}", FormulaRule(formula=['G2="OK"'], fill=DONE_FILL))
    ws.column_dimensions["I"].width = 4
    ws.column_dimensions["J"].width = 42
    ws.column_dimensions["K"].width = 10
    ws["J1"] = "Settings and results"
    ws["J1"].font = Font(bold=True)
    ws["J2"], ws["K2"] = "Same-day handoff counts as a conflict (Yes/No)", "Yes"
    ws["K2"].fill = INPUT_FILL
    res = [
        ("Bookings with a conflict", f'=COUNTIF(G2:G{last},"Conflict")', "0"),
        ("Total bookings", f"=COUNTA(A2:A{last})", "0"),
        ("Rows to fix (Check)", f'=COUNTIF(G2:G{last},"Check")', "0"),
        ("Check: both counts agree", f'=IF(SUM(F2:F{last})=SUM(H2:H{last}),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=3):
        ws[f"J{i}"], ws[f"K{i}"] = label, f
        ws[f"K{i}"].number_format = fmt
        ws[f"K{i}"].fill = FORMULA_FILL
        ws[f"K{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("K6", FormulaRule(formula=['K6="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Find Overlapping Dates and Double Bookings",
        "",
        "1. Bookings tab: type the booking, the space, the client and the start and end dates in A to E (yellow). Rows 2 to 40 have formulas ready. Pick the space from the dropdown in B.",
        "2. F2 counts the other bookings in the same space whose dates overlap this one: =COUNTIFS($B$2:$B$40,B2,$D$2:$D$40,\"<=\"&E2,$E$2:$E$40,\">=\"&D2)-1. Two ranges overlap when each one starts on or before the day the other ends. The -1 removes the row itself.",
        "3. G2 shows Conflict when F2 is above 0, OK when it is 0, and Check when a date is missing, typed as text, or the end date is before the start date.",
        "4. K2 switches how a same-day handoff is treated. Yes means an end date equal to the next start date is a conflict. No means it is allowed (the formula then uses < and > instead of <= and >=).",
        "5. H2 counts the same thing with SUMPRODUCT. K6 shows OK when the two counts agree across all rows.",
        "6. A date typed as text (left aligned, no date format) is ignored by COUNTIFS, so the other bookings do not see it. That row itself shows Check. Retype it as a date.",
        "7. A booking with a conflict is flagged on every row involved. Three bookings that overlap at once show 2 on the middle one.",
        "",
        "Bookings and clients are fictional. This checks dates only: it does not know about leases, deposits or contract terms. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb.save(OUT / "tidy-tabs-overlapping-bookings.xlsx")


def cash_drawer_count():
    wb = Workbook()
    ws = wb.active
    ws.title = "Count"
    header(ws, ["Denomination", "Value", "Count", "Amount"], [20, 11, 9, 13])
    dens = [
        ("$100 bill", 100, 0), ("$50 bill", 50, 0), ("$20 bill", 20, 6), ("$10 bill", 10, 9),
        ("$5 bill", 5, 14), ("$1 bill", 1, 23), ("Quarter", 0.25, 18), ("Dime", 0.10, 12),
        ("Nickel", 0.05, 9), ("Penny", 0.01, 20),
    ]
    for r, (name, val, cnt) in enumerate(dens, start=2):
        ws.append([name, val, cnt])
        ws[f"D{r}"] = f'=IF(C{r}="","",IF(NOT(ISNUMBER(C{r})),"Check",IF(AND(C{r}>=0,C{r}=INT(C{r})),B{r}*C{r},"Check")))'
        ws[f"B{r}"].number_format = USD
        ws[f"D{r}"].number_format = USD
        ws[f"C{r}"].fill = INPUT_FILL
        ws[f"D{r}"].fill = FORMULA_FILL
    ws["A12"], ws["C12"], ws["D12"] = "Total counted", "=SUM(C2:C11)", "=SUM(D2:D11)"
    ws["D12"].number_format = USD
    for a in ("A12", "C12", "D12"):
        ws[a].font = Font(bold=True)
        ws[a].fill = FORMULA_FILL
    ws.conditional_formatting.add("D2:D11", FormulaRule(formula=['D2="Check"'], fill=ALERT_FILL))
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["F"].width = 14
    ws["E1"] = "Closing numbers"
    ws["E1"].font = Font(bold=True)
    ws["E2"], ws["F2"] = "Starting float", 150
    ws["E3"], ws["F3"] = "Cash sales for the day", 161.85
    for a in ("F2", "F3"):
        ws[a].fill = INPUT_FILL
        ws[a].number_format = USD
    res = [
        ("Expected cash", "=F2+F3", USD),
        ("Over (+) or short (-)", "=ROUND(D12-F4,2)", USD),
        ("Status", '=IF(COUNTIF(D2:D11,"Check")>0,"Check counts",IF(F5=0,"Balanced",IF(F5>0,"Over","Short")))', "General"),
        ("Check: SUMPRODUCT total", "=SUMPRODUCT(B2:B11,C2:C11)", USD),
        ("Check: both totals agree", '=IF(ROUND(F7-D12,2)=0,"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=4):
        ws[f"E{i}"], ws[f"F{i}"] = label, f
        ws[f"F{i}"].number_format = fmt
        ws[f"F{i}"].fill = FORMULA_FILL
        ws[f"F{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("F6", FormulaRule(formula=['OR(F6="Over",F6="Short",F6="Check counts")'], fill=ALERT_FILL))
    ws.conditional_formatting.add("F6", FormulaRule(formula=['F6="Balanced"'], fill=DONE_FILL))
    ws.conditional_formatting.add("F8", FormulaRule(formula=['F8="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Cash Drawer Count Over or Short",
        "",
        "1. Count tab: type how many of each bill and coin are in the drawer in C (yellow). A blank count is treated as none.",
        "2. D2: =B2*C2 gives the amount for each row. D12 adds them: =SUM(D2:D11). C12 is the number of bills and coins.",
        "3. F2 is the cash you started the day with (the float). F3 is the cash sales for the day from your own records. F4 adds them to give the cash you expect to find.",
        "4. F5: =ROUND(D12-F4,2) is the difference. ROUND is there because decimals like 0.10 are stored in binary: a difference of one cent can otherwise come out as 0.00999999999999091, which is not equal to 0.01.",
        "5. F6 shows Balanced, Over or Short. A count that is negative, a fraction or typed as text shows Check in D, and F6 says Check counts.",
        "6. F7 totals the drawer a second way with SUMPRODUCT. F8 shows OK when it matches D12.",
        "7. To count another drawer, save a copy of the file or clear column C and F3.",
        "",
        "Counts and sales are fictional. This is arithmetic only, not accounting advice. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-cash-drawer-count.xlsx")


def nth_weekday_of_month():
    wb = Workbook()
    ws = wb.active
    ws.title = "Dates"
    header(ws, ["Item", "Value"], [34, 18])
    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    ws["A2"], ws["B2"] = "Year", 2026
    ws["A3"], ws["B3"] = "Month (1 to 12)", 11
    ws["A4"], ws["B4"] = "Weekday", "Saturday"
    ws["A5"], ws["B5"] = "Which one (1 to 5)", 2
    for a in ("B2", "B3", "B4", "B5"):
        ws[a].fill = INPUT_FILL
        ws[a].alignment = Alignment(horizontal="right")
    ws["A6"], ws["B6"] = "Weekday number (1 = Sunday)", '=IFERROR(MATCH(B4,$E$2:$E$8,0),"Check weekday")'
    ws["A7"] = "Inputs OK?"
    ws["B7"] = ('=IF(AND(ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B5),ISNUMBER(B6)),'
                'IF(AND(B2>=1900,B2<=9999,B2=INT(B2),B3>=1,B3<=12,B3=INT(B3),B5>=1,B5<=5,B5=INT(B5)),"Yes","No"),"No")')
    first = "DATE(B2,B3,1)"
    nth = f"{first}+MOD(B6-WEEKDAY({first},1),7)+7*(B5-1)"
    ws["A8"] = "Date of the Nth weekday"
    ws["B8"] = f'=IF(B7<>"Yes","Check inputs",IF(MONTH({nth})=B3,{nth},"No such date"))'
    eom = f"EOMONTH({first},0)"
    ws["A9"] = "Date of the last such weekday"
    ws["B9"] = f'=IF(B7<>"Yes","Check inputs",{eom}-MOD(WEEKDAY({eom},1)-B6,7))'
    ws["A10"] = "Check: result falls on that weekday"
    ws["B10"] = ('=IF(AND(ISNUMBER(B8),ISNUMBER(B9)),IF(AND(WEEKDAY(B8,1)=B6,WEEKDAY(B9,1)=B6),"OK","CHECK"),'
                 'IF(ISNUMBER(B9),IF(WEEKDAY(B9,1)=B6,"OK","CHECK"),""))')
    for a in ("B6", "B7", "B8", "B9", "B10"):
        ws[a].fill = FORMULA_FILL
        ws[a].alignment = Alignment(horizontal="right")
    ws["B8"].number_format = US_DATE
    ws["B9"].number_format = US_DATE
    ws.conditional_formatting.add("B7:B9", FormulaRule(formula=['OR(B7="No",B7="Check inputs",B7="No such date")'], fill=ALERT_FILL))
    ws.conditional_formatting.add("B10", FormulaRule(formula=['B10="CHECK"'], fill=ALERT_FILL))
    dv = DataValidation(type="list", formula1="=$E$2:$E$8", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B4")
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 14
    ws["E1"] = "Weekday list"
    ws["E1"].font = Font(bold=True)
    for i, d in enumerate(days, start=2):
        ws[f"E{i}"] = d

    sc = wb.create_sheet("Schedule")
    header(sc, ["Month", "Date", "Weekday"], [16, 16, 14])
    ok = ('AND(ISNUMBER(Dates!$B$2),ISNUMBER(Dates!$B$5),ISNUMBER(Dates!$B$6),Dates!$B$2>=1900,Dates!$B$2<=9999,'
          'Dates!$B$2=INT(Dates!$B$2),Dates!$B$5>=1,Dates!$B$5<=5,Dates!$B$5=INT(Dates!$B$5))')
    for r in range(2, 14):
        sc[f"A{r}"] = f"=IF({ok},DATE(Dates!$B$2,1,1),\"\")" if r == 2 else f'=IF(A{r-1}="","",EDATE(A{r-1},1))'
        x = f"A{r}+MOD(Dates!$B$6-WEEKDAY(A{r},1),7)+7*(Dates!$B$5-1)"
        sc[f"B{r}"] = f'=IF(A{r}="","",IF(MONTH({x})=MONTH(A{r}),{x},"No such date"))'
        sc[f"C{r}"] = f'=IF(ISNUMBER(B{r}),Dates!$B$4,"")'
        sc[f"A{r}"].number_format = "mmmm yyyy"
        sc[f"B{r}"].number_format = US_DATE
        for col in "ABC":
            sc[f"{col}{r}"].fill = FORMULA_FILL
        sc[f"B{r}"].alignment = Alignment(horizontal="right")
    sc.conditional_formatting.add("B2:B13", FormulaRule(formula=['B2="No such date"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Date of the Nth Weekday of a Month",
        "",
        "1. Dates tab: type the year, the month number, pick the weekday from the dropdown and type which one you want (1 to 5) in B2 to B5 (yellow).",
        "2. B6 turns the weekday name into a number with MATCH (1 = Sunday to 7 = Saturday, the WEEKDAY default).",
        "3. B8: =DATE(B2,B3,1)+MOD(B6-WEEKDAY(DATE(B2,B3,1),1),7)+7*(B5-1). The first part finds the first wanted weekday on or after the 1st, the last part adds 7 days for each extra week.",
        "4. B8 is wrapped in a MONTH check. If the result lands in the next month (for example a 5th Saturday that does not exist), it shows No such date.",
        "5. B9 is the last such weekday of the month: =EOMONTH(DATE(B2,B3,1),0)-MOD(WEEKDAY(EOMONTH(DATE(B2,B3,1),0),1)-B6,7). It steps back from the last day of the month.",
        "6. A month typed as 13 would roll into the next year inside DATE, so B7 checks the inputs first and the results show Check inputs. A year typed as text does the same.",
        "7. Schedule tab: the same weekday and number for all 12 months of the year in B2. Change the yellow cells on Dates and it follows.",
        "",
        "This is date math only. It does not know about holidays or market rules. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-nth-weekday-of-month.xlsx")


def format_phone_numbers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Phones"
    header(ws, ["Phone as typed", "Digits only", "Without leading 1", "Formatted"], [24, 16, 18, 18])
    raw = ["(212) 555-0101", "212.555.0102", "+1 212-555-0103", 2125550104, "212 555 0105",
           "1-212-555-0106", "(415)555-0107", "415.555.0108", "212-555-010",
           "212-555-0111 x12", "21255501130", "call after 5"]
    last = 40
    for r in range(2, last + 1):
        ws.append([raw[r - 2]] if r - 2 < len(raw) else [])
        s = f'A{r}'
        for ch in (" ", "-", "(", ")", ".", "+"):
            s = f'SUBSTITUTE({s},"{ch}","")'
        ws[f"B{r}"] = f'=IF(A{r}="","",{s})'
        ws[f"C{r}"] = f'=IF(B{r}="","",IF(AND(LEN(B{r})=11,LEFT(B{r},1)="1"),MID(B{r},2,10),B{r}))'
        digits = f'SUMPRODUCT(--ISNUMBER(FIND(MID(C{r},{{1,2,3,4,5,6,7,8,9,10}},1),"0123456789")))=10'
        ws[f"D{r}"] = (f'=IF(A{r}="","",IF(AND(LEN(C{r})=10,{digits}),'
                       f'"("&LEFT(C{r},3)&") "&MID(C{r},4,3)&"-"&RIGHT(C{r},4),"Check"))')
        for col in "ABC":
            ws[f"{col}{r}"].number_format = "@"
        ws[f"A{r}"].fill = INPUT_FILL
        for col in "BCD":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"D2:D{last}", FormulaRule(formula=['D2="Check"'], fill=ALERT_FILL))
    ws.column_dimensions["E"].width = 4
    ws.column_dimensions["F"].width = 30
    ws.column_dimensions["G"].width = 10
    ws["F1"] = "Results"
    ws["F1"].font = Font(bold=True)
    res = [
        ("Rows with a phone entry", f"=COUNTA(A2:A{last})", "0"),
        ("Formatted", f"=SUMPRODUCT(--(LEN(D2:D{last})=14))", "0"),
        ("Rows to check", f'=COUNTIF(D2:D{last},"Check")', "0"),
        ("Check: formatted + check = rows", f'=IF(G3+G4=G2,"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=2):
        ws[f"F{i}"], ws[f"G{i}"] = label, f
        ws[f"G{i}"].number_format = fmt
        ws[f"G{i}"].fill = FORMULA_FILL
        ws[f"G{i}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add("G5", FormulaRule(formula=['G5="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Format Phone Numbers as (555) 123-4567",
        "",
        "1. Phones tab: paste or type the phone numbers in A (yellow). Rows 2 to 40 have formulas ready. Columns A to C are formatted as text so a leading zero is kept.",
        "2. B2 strips the spaces, hyphens, parentheses, dots and plus signs with nested SUBSTITUTE: =SUBSTITUTE(SUBSTITUTE(...(A2,\" \",\"\"),\"-\",\"\"),...).",
        "3. C2: =IF(AND(LEN(B2)=11,LEFT(B2,1)=\"1\"),MID(B2,2,10),B2) drops a leading 1 (the US country code) from an 11 digit number.",
        "4. D2 rebuilds a 10 digit number as (555) 123-4567 with LEFT, MID and RIGHT. Anything else shows Check: too few or too many digits, letters, or an extension such as x12.",
        "5. G2 to G4 count the rows with an entry, the formatted rows and the Check rows. G5 shows OK when formatted plus Check equals the rows with an entry.",
        "6. Copy D and use Paste values only when you need the text without formulas, for example before importing contacts or printing labels.",
        "7. This only fixes the format. It does not check that a number really exists or can be called. Extensions and numbers from other countries are left as Check.",
        "",
        "Phone numbers are fictional (555-01xx style). Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-format-phone-numbers.xlsx")


def scale_recipe_batch():
    wb = Workbook()
    ws = wb.active
    ws.title = "Recipe"
    for col, w in zip("ABCDE", [30, 14, 18, 14, 18]):
        ws.column_dimensions[col].width = w
    ws["A1"], ws["C1"] = "Base batch yield (jars)", 12
    ws["A2"], ws["C2"] = "Target yield (jars)", 30
    ws["A3"], ws["C3"] = "Scale factor", '=IF(AND(ISNUMBER(C1),C1>0),IF(ISNUMBER(C2),C2/C1,0),"")'
    ws["A4"], ws["C4"] = "Check", ('=IF(C3="","Check base yield",IF(C3=0,"Check target yield",'
                                  'IF(SUMPRODUCT((A6:A30<>"")*(1-ISNUMBER(B6:B30)))>0,"Check quantities","OK")))')
    for a in ("C1", "C2"):
        ws[a].fill = INPUT_FILL
    for a in ("C3", "C4"):
        ws[a].fill = FORMULA_FILL
    ws["C3"].number_format = '0.00"x"'
    for a in ("C1", "C2", "C3", "C4"):
        ws[a].alignment = Alignment(horizontal="right")
    for c in ("A1", "A2", "A3", "A4"):
        ws[c].font = Font(bold=True)
    names = ["Ingredient", "Base quantity", "Unit", "Scaled quantity", "Buy whole units"]
    for col, n in zip("ABCDE", names):
        c = ws[f"{col}5"]
        c.value, c.fill, c.font = n, HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    items = [
        ("Strawberries", 64, "oz"), ("Granulated sugar", 6, "cup"), ("Lemon juice", 0.25, "cup"),
        ("Powdered pectin", 4, "tsp"), ("Lemons", 3, "each"), ("Vanilla extract", 1.5, "tsp"),
        ("Salt", 0.25, "tsp"), ("Jar lids", 12, "each"),
    ]
    last = 30
    for r in range(6, last + 1):
        if r - 6 < len(items):
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = items[r - 6]
        ws[f"D{r}"] = f'=IF(OR(C$3="",NOT(ISNUMBER(B{r}))),"",ROUND(B{r}*$C$3,2))'
        ws[f"E{r}"] = (f'=IF(D{r}="","",IF(C{r}="each",ROUNDUP(ROUND(B{r}*$C$3,6),0),ROUND(B{r}*$C$3,2)))')
        ws[f"B{r}"].number_format = "0.##"
        ws[f"D{r}"].number_format = "0.00"
        ws[f"E{r}"].number_format = "0.00"
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    dv = DataValidation(type="list", formula1='"oz,cup,tsp,each"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C6:C{last}")
    ws.conditional_formatting.add("C4", FormulaRule(formula=['C4<>"OK"'], fill=ALERT_FILL))
    ws.freeze_panes = "A6"

    notes_sheet(wb, [
        "Tidy Tabs: Scale a Recipe or Batch Up or Down",
        "",
        "1. Recipe tab: type how many jars (or loaves, bars, candles) the recipe makes in C1 and how many you want in C2 (yellow).",
        "2. C3: =C2/C1 is the scale factor, shown as a multiplier. 30 jars from a 12 jar recipe is 2.50x.",
        "3. Type each ingredient, its base quantity and its unit in A to C. Write fractions as decimals: 1/3 cup is 0.33 and 1/4 cup is 0.25.",
        "4. D6: =ROUND(B6*$C$3,2). The $ signs lock the factor to C3, so the formula can be copied down and still points at the same cell. B6 has no $ and moves to B7, B8 and so on.",
        "5. E6 is for things you cannot buy in parts. When the unit is each, it rounds up to a whole number with ROUNDUP (3 lemons times 2.5 is 7.5, so 8). Other units are rounded to 2 decimals like D.",
        "6. If C2 is blank or 0 the factor is 0 and every quantity becomes 0: C4 says Check target yield. If C1 is 0 or blank, C3 stays empty and C4 says Check base yield. A quantity typed as text is left out and C4 says Check quantities.",
        "7. Rows 6 to 30 have formulas ready; copy the formula cells down for more.",
        "",
        "The jam recipe and quantities are fictional. Cooking times, baking chemistry, set points and food safety can change at larger sizes, and this sheet does not cover them. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Recipe"].print_title_rows = "5:5"
    wb.save(OUT / "tidy-tabs-scale-recipe-batch.xlsx")


def lookup_two_criteria():
    wb = Workbook()
    pr = wb.active
    pr.title = "Prices"
    header(pr, ["Size", "Finish", "Price", "Key (Size|Finish)", "Row check"], [12, 12, 11, 20, 20])
    rows = [
        ("Small", "Matte", 12.00), ("Small", "Gloss", 14.00), ("Small", "Canvas", 32.00),
        ("Medium", "Matte", 18.00), ("Medium", "Gloss", 21.00), ("Medium", "Canvas", 45.00),
        ("Large", "Matte", 26.00), ("Large", "Gloss", 30.00), ("Large", "Canvas", 64.00),
    ]
    last = 30
    for r in range(2, last + 1):
        if r - 2 < len(rows):
            pr[f"A{r}"], pr[f"B{r}"], pr[f"C{r}"] = rows[r - 2]
        pr[f"D{r}"] = f'=IF(A{r}="","",A{r}&"|"&B{r})'
        pr[f"E{r}"] = (f'=IF(A{r}="","",IF(NOT(ISNUMBER(C{r})),"Price not a number",'
                       f'IF(COUNTIFS($A$2:$A${last},A{r},$B$2:$B${last},B{r})>1,"Duplicate","OK")))')
        pr[f"C{r}"].number_format = USD
        for col in "ABC":
            pr[f"{col}{r}"].fill = INPUT_FILL
        for col in "DE":
            pr[f"{col}{r}"].fill = FORMULA_FILL
    pr.conditional_formatting.add(f"E2:E{last}", FormulaRule(formula=['AND(E2<>"",E2<>"OK")'], fill=ALERT_FILL))

    lk = wb.create_sheet("Lookup", 1)
    lk.column_dimensions["A"].width = 46
    lk.column_dimensions["B"].width = 20
    p = f"Prices!$C$2:$C${last}"
    a = f"Prices!$A$2:$A${last}"
    b = f"Prices!$B$2:$B${last}"
    k = f"Prices!$D$2:$D${last}"
    blank = 'OR($B$1="",$B$2="")'
    cells = [
        ("A1", "Size", "B1", "Large"),
        ("A2", "Finish", "B2", "Gloss"),
        ("A4", "Price: INDEX/MATCH on the Key column", "B4",
         f'=IF({blank},"",IFERROR(INDEX({p},MATCH(B1&"|"&B2,{k},0)),"Not on price list"))'),
        ("A5", "Price: SUMIFS (numbers only, 0 if no match)", "B5",
         f'=IF({blank},"",SUMIFS({p},{a},B1,{b},B2))'),
        ("A6", "Price: INDEX/MATCH without a helper column", "B6",
         f'=IF({blank},"",IFERROR(INDEX({p},MATCH(B1&"|"&B2,INDEX({a}&"|"&{b},0),0)),"Not on price list"))'),
        ("A7", "Rows on the price list with this size and finish", "B7",
         f'=IF({blank},"",COUNTIFS({a},B1,{b},B2))'),
        ("A8", "Check: the three prices agree", "B8",
         f'=IF({blank},"Pick a size and finish",IF(B7=0,"Not on price list",IF(B7>1,"Duplicate rows",'
         'IF(AND(B4=B5,B6=B5),"OK","CHECK"))))'),
    ]
    for la, label, va, val in cells:
        lk[la], lk[va] = label, val
        lk[la].font = Font(bold=True)
        lk[va].alignment = Alignment(horizontal="right")
        lk[va].fill = INPUT_FILL if va in ("B1", "B2") else FORMULA_FILL
    for va in ("B4", "B5", "B6"):
        lk[va].number_format = USD
    for va, lst in (("B1", '"Small,Medium,Large"'), ("B2", '"Matte,Gloss,Canvas"')):
        dv = DataValidation(type="list", formula1=lst, allow_blank=True)
        lk.add_data_validation(dv)
        dv.add(va)
    lk.conditional_formatting.add("B8", FormulaRule(formula=['B8<>"OK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Look Up a Price With Two Criteria (Size and Finish)",
        "",
        "1. Prices tab: one row per size and finish with its price in A to C (yellow). Rows 2 to 30 have formulas ready.",
        "2. Prices!D2 builds a Key that joins the two criteria: =A2&\"|\"&B2 gives Large|Gloss. The | keeps pairs such as \"Large\" + \"Gloss\" apart from other text that would run together.",
        "3. Lookup tab: pick a size in B1 and a finish in B2 (yellow dropdowns).",
        "4. B4: =INDEX(Prices!C2:C30,MATCH(B1&\"|\"&B2,Prices!D2:D30,0)) finds the Key and returns the price on that row.",
        "5. B5: =SUMIFS(Prices!C2:C30,Prices!A2:A30,B1,Prices!B2:B30,B2) adds the prices of every row that matches both. It works only for numbers and shows 0 when nothing matches.",
        "6. B6 does the same as B4 without the Key column: INDEX(Prices!A2:A30&\"|\"&Prices!B2:B30,0) joins the two columns inside the formula, so no Ctrl+Shift+Enter is needed.",
        "7. B7 counts how many rows match both criteria. B8 shows OK when the three prices agree. It shows Not on price list when no row matches, Duplicate rows when the same pair is listed twice (SUMIFS adds both, MATCH takes the first), and CHECK when a price is typed as text.",
        "8. Matching ignores upper and lower case (large finds Large), but an extra space does not match: \"Large \" with a trailing space shows Not on price list.",
        "9. Lookup only. The prices are fictional and are not pricing advice.",
        "",
        "Sample shop and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-lookup-two-criteria.xlsx")


def stock_on_hand_log():
    wb = Workbook()
    log = wb.active
    log.title = "Log"
    header(log, ["Date", "SKU", "Type", "Qty", "Note", "Row check"], [12, 12, 13, 8, 30, 11])
    entries = [
        ((9, 1), "SOAP-OAT", "Received", 48, "Cured batch"),
        ((9, 1), "SOAP-LAV", "Received", 36, "Cured batch"),
        ((9, 2), "CNDL-CED", "Received", 24, "Poured batch"),
        ((9, 2), "CNDL-VAN", "Received", 24, "Poured batch"),
        ((9, 3), "BALM-LIP", "Received", 60, "Filled tubes"),
        ((9, 5), "SOAP-OAT", "Sold", 14, "Saturday farmers market"),
        ((9, 5), "CNDL-CED", "Sold", 6, "Saturday farmers market"),
        ((9, 5), "BALM-LIP", "Sold", 18, "Saturday farmers market"),
        ((9, 8), "SOAP-LAV", "Sold", 9, "Etsy orders"),
        ((9, 10), "CNDL-VAN", "Sold", 7, "Etsy orders"),
        ((9, 12), "SOAP-OAT", "Sold", 11, "Saturday farmers market"),
        ((9, 12), "SOAP-LAV", "Sold", 8, "Saturday farmers market"),
        ((9, 12), "BALM-LIP", "Sold", 15, "Saturday farmers market"),
        ((9, 14), "CNDL-CED", "Adjustment", -1, "Cracked jar"),
        ((9, 15), "SOAP-OAT", "Received", 24, "Cured batch"),
        ((9, 17), "CNDL-VAN", "Sold", 9, "Wholesale order, gift shop"),
        ((9, 19), "SOAP-OAT", "Sold", 16, "Saturday farmers market"),
        ((9, 19), "CNDL-CED", "Sold", 8, "Saturday farmers market"),
        ((9, 19), "BALM-LIP", "Sold", 12, "Saturday farmers market"),
        ((9, 22), "SOAP-LAV", "Adjustment", -2, "Recount"),
        ((9, 24), "CNDL-CED", "Received", 12, "Poured batch"),
        ((9, 26), "SOAP-OAT", "Sold", 13, "Saturday farmers market"),
        ((9, 26), "SOAP-LAV", "Sold", 10, "Saturday farmers market"),
        ((9, 26), "CNDL-VAN", "Sold", 5, "Saturday farmers market"),
        ((9, 29), "BALM-LIP", "Adjustment", 3, "Recount found a box"),
        ((10, 1), "BALM-LIP", "Sold", 9, "Etsy orders"),
        ((10, 3), "SOAP-OAT", "Sold", 12, "Saturday farmers market"),
        ((10, 3), "CNDL-CED", "Sold", 10, "Saturday farmers market"),
        ((10, 3), "SOAP-LAV", "Sold", 4, "Saturday farmers market"),
        ((10, 5), "CNDL-VAN", "Received", 12, "Poured batch"),
    ]
    llast = 500
    slast = 20
    for r in range(2, llast + 1):
        if r - 2 < len(entries):
            (m, d), sku, typ, qty, note = entries[r - 2]
            log.append([date(2026, m, d), sku, typ, qty, note])
        log[f"F{r}"] = (f'=IF(AND(A{r}="",B{r}="",C{r}="",D{r}=""),"",'
                        f'IF(OR(B{r}="",NOT(ISNUMBER(D{r})),COUNTIF(Stock!$A$2:$A${slast},B{r})=0,'
                        f'AND(C{r}<>"Received",C{r}<>"Sold",C{r}<>"Adjustment")),"Check","OK"))')
        log[f"A{r}"].number_format = US_DATE
        if r <= 60:
            for col in "ABCDE":
                log[f"{col}{r}"].fill = INPUT_FILL
            log[f"F{r}"].fill = FORMULA_FILL
    types = DataValidation(type="list", formula1='"Received,Sold,Adjustment"', allow_blank=True)
    log.add_data_validation(types)
    types.add(f"C2:C{llast}")
    skus = DataValidation(type="list", formula1=f"=Stock!$A$2:$A${slast}", allow_blank=True)
    log.add_data_validation(skus)
    skus.add(f"B2:B{llast}")
    log.conditional_formatting.add(f"F2:F{llast}", FormulaRule(formula=['F2="Check"'], fill=ALERT_FILL))

    st = wb.create_sheet("Stock", 1)
    header(st, ["SKU", "Received", "Sold", "Adjustments", "On hand", "Minimum", "Status"],
           [12, 10, 9, 12, 10, 10, 10])
    items = [("SOAP-OAT", 10), ("SOAP-LAV", 10), ("CNDL-CED", 6), ("CNDL-VAN", 6), ("BALM-LIP", 8)]
    lb, lc, ld = (f"Log!${c}$2:${c}${llast}" for c in "BCD")
    for r in range(2, slast + 1):
        if r - 2 < len(items):
            st[f"A{r}"], st[f"F{r}"] = items[r - 2]
        st[f"B{r}"] = f'=IF(A{r}="","",SUMIFS({ld},{lb},A{r},{lc},"Received"))'
        st[f"C{r}"] = f'=IF(A{r}="","",SUMIFS({ld},{lb},A{r},{lc},"Sold"))'
        st[f"D{r}"] = f'=IF(A{r}="","",SUMIFS({ld},{lb},A{r},{lc},"Adjustment"))'
        st[f"E{r}"] = f'=IF(A{r}="","",B{r}-C{r}+D{r})'
        st[f"G{r}"] = (f'=IF(A{r}="","",IF(E{r}<0,"Check",IF(NOT(ISNUMBER(F{r})),"OK",'
                       f'IF(E{r}<=F{r},"Reorder","OK"))))')
        for col in "AF":
            st[f"{col}{r}"].fill = INPUT_FILL
        for col in "BCDEG":
            st[f"{col}{r}"].fill = FORMULA_FILL
    st.conditional_formatting.add(f"G2:G{slast}", FormulaRule(formula=['G2="Reorder"'], fill=ALERT_FILL))
    st.conditional_formatting.add(f"G2:G{slast}", FormulaRule(formula=['G2="Check"'], fill=ALERT_FILL))
    st.column_dimensions["H"].width = 4
    st.column_dimensions["I"].width = 40
    st.column_dimensions["J"].width = 11
    st["I1"] = "Totals and check"
    st["I1"].font = Font(bold=True)
    res = [
        ("Received in the whole log", f'=SUMIFS({ld},{lc},"Received")'),
        ("Sold in the whole log", f'=SUMIFS({ld},{lc},"Sold")'),
        ("Adjustments in the whole log", f'=SUMIFS({ld},{lc},"Adjustment")'),
        ("Log total (received - sold + adjustments)", "=J2-J3+J4"),
        ("On hand, all SKUs on this tab", f"=SUM(E2:E{slast})"),
        ("Log rows to check", f'=COUNTIF(Log!$F$2:$F${llast},"Check")'),
        ("SKUs to reorder", f'=COUNTIF(G2:G{slast},"Reorder")'),
        ("Check: on hand = log total, nothing to check",
         f'=IF(AND(J6=J5,J7=0,COUNTIF(G2:G{slast},"Check")=0),"OK","CHECK")'),
    ]
    for i, (label, f) in enumerate(res, start=2):
        st[f"I{i}"], st[f"J{i}"] = label, f
        st[f"J{i}"].fill = FORMULA_FILL
        st[f"J{i}"].alignment = Alignment(horizontal="right")
    st.conditional_formatting.add("J9", FormulaRule(formula=['J9="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Stock on Hand From a Log of Items In and Out",
        "",
        "1. Stock tab: list each SKU once in A and the count you want to reorder at in F (yellow). Rows 2 to 20 have formulas ready.",
        "2. Log tab: add one row every time stock moves. Date, SKU (dropdown from the Stock tab), Type (Received, Sold or Adjustment), Qty and a note.",
        "3. Type Received and Sold as positive numbers. Type an Adjustment with its sign: -2 for two broken bars, 3 for three found in a recount.",
        "4. Stock!B2: =SUMIFS(Log!D2:D500,Log!B2:B500,A2,Log!C2:C500,\"Received\") adds every Received row for that SKU. C2 and D2 do the same for Sold and Adjustment.",
        "5. On hand E2: =B2-C2+D2. Status G2 shows Reorder when On hand is at or below the Minimum and Check when On hand is below zero (more sold than received).",
        "6. Log!F shows Check for a row with a blank SKU, a SKU that is not on the Stock tab (watch for a trailing space), a Type outside the list or a Qty typed as text. Those rows are left out of the Stock tab.",
        "7. Stock!J9 shows OK when the on hand total equals received minus sold plus adjustments for the whole log, no log row needs checking and no SKU is below zero.",
        "8. SKU matching ignores upper and lower case (soap-oat counts as SOAP-OAT). The Date column is not used in the totals; it keeps the history.",
        "9. Counts only. This sheet does not value inventory or work out cost of goods sold.",
        "10. The print area covers the first 40 log rows. Change the print area to print a longer log.",
        "",
        "Shop, SKUs and quantities are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Log"].print_area = "A1:F41"
    wb["Stock"].print_area = f"A1:J{slast + 1}"
    wb.save(OUT / "tidy-tabs-stock-on-hand-log.xlsx")


def email_domain_split():
    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"
    header(ws, ["Email as typed", "Username", "Domain", "Domain (lower case)", "Shape"], [30, 18, 22, 22, 9])
    emails = [
        "Jane.Doe@Example.com", "sam@shop.example.org", "rosa.m+orders@example.net", "lee_k@example.com",
        "BILL.R@EXAMPLE.ORG", "maria@@example.com", "tomexample.com", "ann lee@example.net",
        "kim@example", "pat@example.com ", "dev.team@mail.example.com", "hello@example.org",
        None, "orders@example.net", "chris.p@example.com",
    ]
    last = 200
    for r in range(2, last + 1):
        e = emails[r - 2] if r - 2 < len(emails) else None
        ws.append([e])
        ws[f"B{r}"] = f'=IF(A{r}="","",IFERROR(LEFT(A{r},FIND("@",A{r})-1),""))'
        ws[f"C{r}"] = f'=IF(A{r}="","",IFERROR(MID(A{r},FIND("@",A{r})+1,LEN(A{r})),""))'
        ws[f"D{r}"] = f'=IF(C{r}="","",LOWER(C{r}))'
        ws[f"E{r}"] = (f'=IF(A{r}="","",IF(AND(LEN(A{r})-LEN(SUBSTITUTE(A{r},"@",""))=1,LEFT(A{r},1)<>"@",'
                       f'ISNUMBER(FIND(".",A{r},FIND("@",A{r})+2)),RIGHT(A{r},1)<>".",'
                       f'ISERROR(FIND(" ",A{r}))),"OK","Check"))')
        ws[f"A{r}"].number_format = "@"
        if r <= 60:
            ws[f"A{r}"].fill = INPUT_FILL
            for col in "BCDE":
                ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"E2:E{last}", FormulaRule(formula=['E2="Check"'], fill=ALERT_FILL))

    sm = wb.create_sheet("Summary", 1)
    header(sm, ["Domain", "OK rows"], [26, 10])
    doms = ["example.com", "example.org", "example.net"]
    d = f"Emails!$D$2:$D${last}"
    e = f"Emails!$E$2:$E${last}"
    for r in range(2, 8):
        if r - 2 < len(doms):
            sm[f"A{r}"] = doms[r - 2]
        sm[f"B{r}"] = f'=IF(A{r}="","",COUNTIFS({d},TRIM(A{r}),{e},"OK"))'
        sm[f"A{r}"].fill = INPUT_FILL
        sm[f"B{r}"].fill = FORMULA_FILL
    sm["A8"], sm["B8"] = "Other domains", "=B11-SUM(B2:B7)"
    sm["A10"], sm["B10"] = "Rows with an entry", f"=COUNTA(Emails!A2:A{last})"
    sm["A11"], sm["B11"] = "OK rows", f'=COUNTIF({e},"OK")'
    sm["A12"], sm["B12"] = "Rows to check", f'=COUNTIF({e},"Check")'
    sm["A13"], sm["B13"] = "Check: OK + Check = rows", '=IF(B11+B12=B10,"OK","CHECK")'
    for a in ("A8", "A10", "A11", "A12", "A13"):
        sm[a].font = Font(bold=True)
    for a in ("B8", "B10", "B11", "B12", "B13"):
        sm[a].fill = FORMULA_FILL
        sm[a].alignment = Alignment(horizontal="right")
    sm.conditional_formatting.add("B13", FormulaRule(formula=['B13="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Split an Email Address Into Username and Domain",
        "",
        "1. Emails tab: paste the addresses in A (yellow). Rows 2 to 200 have formulas ready.",
        "2. Username B2: =IFERROR(LEFT(A2,FIND(\"@\",A2)-1),\"\") keeps everything before the @.",
        "3. Domain C2: =IFERROR(MID(A2,FIND(\"@\",A2)+1,LEN(A2)),\"\") keeps everything after the @. D2 is the same domain in lower case with LOWER, so Example.com and example.com count together.",
        "4. Shape E2 shows OK when the address has exactly one @, something before it, a dot at least one character after it, no dot at the very end and no spaces. Anything else shows Check.",
        "5. A trailing space after the address shows Check. Remove it with TRIM or retype the cell. A plus sign in the username (name+orders@...) stays OK.",
        "6. Summary tab: type the domains you want to count in A2 to A7 (yellow). B counts the OK rows with that domain. Other domains is every other OK row.",
        "7. Summary!B13 shows OK when OK rows plus Check rows equal the rows with an entry.",
        "8. This checks the shape of the text only. It does not tell you whether a mailbox exists or whether a person agreed to get email from you.",
        "9. Count each domain once on the Summary tab; a domain typed twice is counted twice and Other domains goes below zero. The print area covers the first 40 rows of the Emails tab.",
        "",
        "All addresses are fictional and use example.com, example.org and example.net. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Emails"].print_area = "A1:E41"
    wb.save(OUT / "tidy-tabs-email-domain-split.xlsx")


def on_time_shipping_rate():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order ID", "Order date", "Ship-by date", "Shipped date", "Days late", "Status"],
           [11, 12, 13, 13, 10, 13])
    # Shipped date relative to the ship-by date (negative = early); None = not shipped yet.
    offsets = [-2, -1, 0, 0, 1, -1, -3, 0, 2, -1, 0, -2, -1, 4, 0,
               -1, -2, 0, 1, -1, 0, -3, -1, 0, 3, -2, -1, 0, None, None]
    last = 200
    for r in range(2, last + 1):
        i = r - 2
        if i < len(offsets):
            od = date(2026, 9, 1 + i)
            sb = date.fromordinal(od.toordinal() + 3)
            sh = None if offsets[i] is None else date.fromordinal(sb.toordinal() + offsets[i])
            ws.append([f"#{2401 + i}", od, sb, sh])
        ws[f"F{r}"] = (f'=IF(A{r}="","",IF(OR(NOT(ISNUMBER(B{r})),NOT(ISNUMBER(C{r}))),"Check",'
                       f'IF(D{r}="","Not shipped",IF(NOT(ISNUMBER(D{r})),"Check",'
                       f'IF(OR(D{r}<B{r},C{r}<B{r}),"Check",IF(D{r}<=C{r},"On time","Late"))))))')
        ws[f"E{r}"] = f'=IF(F{r}="Late",D{r}-C{r},IF(F{r}="On time",0,""))'
        for col in "BCD":
            ws[f"{col}{r}"].number_format = US_DATE
        ws[f"E{r}"].number_format = "0"
        if r <= 60:
            for col in "ABCD":
                ws[f"{col}{r}"].fill = INPUT_FILL
            for col in "EF":
                ws[f"{col}{r}"].fill = FORMULA_FILL
    ws.conditional_formatting.add(f"A2:F{last}", FormulaRule(formula=['$F2="Late"'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"F2:F{last}", FormulaRule(formula=['$F2="Check"'], fill=ALERT_FILL))

    sm = wb.create_sheet("Summary", 1)
    sm.column_dimensions["A"].width = 46
    sm.column_dimensions["B"].width = 12
    s = f"Orders!$F$2:$F${last}"
    res = [
        ("Orders listed", f"=COUNTA(Orders!A2:A{last})", "0"),
        ("Orders with a shipped date", f"=COUNT(Orders!D2:D{last})", "0"),
        ("Shipped on time (on or before ship-by)", f'=COUNTIF({s},"On time")', "0"),
        ("Shipped late", f'=COUNTIF({s},"Late")', "0"),
        ("Not shipped yet (left out of the rate)", f'=COUNTIF({s},"Not shipped")', "0"),
        ("Rows to check", f'=COUNTIF({s},"Check")', "0"),
        ("On-time shipping rate", '=IF(B4+B5=0,"",B4/(B4+B5))', "0.0%"),
        ("Average days late (late orders only)", f'=IFERROR(AVERAGEIF(Orders!E2:E{last},">0"),0)', "0.0"),
        ("Most days late", f"=MAX(Orders!E2:E{last})", "0"),
        ("Check: on time + late = shipped, no rows to check", '=IF(AND(B4+B5=B3,B7=0),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=2):
        sm[f"A{i}"], sm[f"B{i}"] = label, f
        sm[f"A{i}"].font = Font(bold=True)
        sm[f"B{i}"].number_format = fmt
        sm[f"B{i}"].fill = FORMULA_FILL
        sm[f"B{i}"].alignment = Alignment(horizontal="right")
    sm["A1"] = "Shipping summary"
    sm["A1"].font = Font(bold=True, size=13)
    sm.conditional_formatting.add("B11", FormulaRule(formula=['B11="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: On-Time Shipping Rate From Ship-By and Shipped Dates",
        "",
        "1. Orders tab: paste or type Order ID, Order date, Ship-by date and Shipped date in A to D (yellow). Rows 2 to 200 have formulas ready. Leave Shipped date blank for orders not shipped yet.",
        "2. Status F2 compares the dates: On time when Shipped date is on or before Ship-by date, Late when it is after, Not shipped when it is blank.",
        "3. F2 shows Check when a date is typed as text, or when the shipped or ship-by date is before the order date. Check rows stay out of the counts.",
        "4. Days late E2 is Shipped date minus Ship-by date for late orders and 0 for on-time orders. Late rows are colored.",
        "5. Summary tab: B8 is the on-time rate, on time divided by (on time + late). Orders not shipped yet are left out. With no shipped orders the rate stays blank instead of #DIV/0!.",
        "6. B9 is the average days late among late orders (0 when none are late) and B10 is the longest delay.",
        "7. B11 shows OK when on time plus late equals the orders with a shipped date and no row needs checking.",
        "8. Dates must be real dates. A date typed as text (usually left-aligned) shows Check. Times of day are not used.",
        "9. Marketplaces set and measure their own shipping targets. Use their seller help pages for the official numbers; this sheet only compares your own dates.",
        "10. The print area covers the first 40 orders. Change the print area to print more.",
        "",
        "Orders and dates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Orders"].print_area = "A1:F41"
    wb.save(OUT / "tidy-tabs-on-time-shipping-rate.xlsx")


def price_per_square_foot():
    wb = Workbook()
    ws = wb.active
    ws.title = "Quote"
    for col, w in zip("ABCDEFGHIJK", [30, 12, 12, 8, 13, 12, 11, 18, 4, 30, 12]):
        ws.column_dimensions[col].width = w
    settings = [
        ("A1", "Rate per sq ft", "B1", 2.50, USD),
        ("A2", "Minimum charge per line", "B2", 25.00, USD),
        ("A3", "Setup fee (once per quote)", "B3", 15.00, USD),
        ("A4", "Settings OK?", "B4",
         '=IF(AND(ISNUMBER(B1),ISNUMBER(B2),ISNUMBER(B3)),IF(AND(B1>0,B2>=0,B3>=0),"OK","Check settings"),"Check settings")',
         "General"),
    ]
    for la, label, va, val, fmt in settings:
        ws[la], ws[va] = label, val
        ws[la].font = Font(bold=True)
        ws[va].number_format = fmt
        ws[va].alignment = Alignment(horizontal="right")
        ws[va].fill = FORMULA_FILL if va == "B4" else INPUT_FILL
    names = ["Job", "Width (in)", "Height (in)", "Qty", "Area per piece (sq ft)", "Line price",
             "Minimum applied", "Size check"]
    for col, n in zip("ABCDEFGH", names):
        c = ws[f"{col}5"]
        c.value, c.fill, c.font = n, HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[5].height = 30
    jobs = [
        ("Vinyl banner", 36, 72, 1), ("Yard signs", 18, 24, 10), ("Window decals", 12, 12, 2),
        ("Poster", 24, 36, 1), ("Canvas prints", 30, 40, 2), ("Table banner", 24, 60, 1),
        ("Door signs", 11, 17, 3), ("Feather flags", 30, 96, 2),
    ]
    first, last = 6, 25
    for r in range(first, last + 1):
        if r - first < len(jobs):
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = jobs[r - first]
        ws[f"H{r}"] = (f'=IF(A{r}&B{r}&C{r}&D{r}="","",IF(OR(NOT(ISNUMBER(B{r})),NOT(ISNUMBER(C{r})),'
                       f'NOT(ISNUMBER(D{r}))),"Check",IF(OR(B{r}<=0,C{r}<=0,D{r}<=0,D{r}<>INT(D{r})),"Check",'
                       f'IF(B{r}*C{r}<36,"Under 36 sq in: feet?","OK"))))')
        ws[f"E{r}"] = f'=IF(OR(H{r}="",H{r}="Check"),"",ROUND(B{r}*C{r}/144,2))'
        ws[f"F{r}"] = f'=IF(OR(E{r}="",$B$4<>"OK"),"",MAX($B$2,ROUND(E{r}*D{r}*$B$1,2)))'
        ws[f"G{r}"] = f'=IF(F{r}="","",IF(ROUND(E{r}*D{r}*$B$1,2)<$B$2,"Yes","No"))'
        ws[f"E{r}"].number_format = "0.00"
        ws[f"F{r}"].number_format = USD
        for col in "ABCD":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "EFGH":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    rg = lambda c: f"{c}{first}:{c}{last}"
    ws.conditional_formatting.add(rg("H"), FormulaRule(formula=[f'AND(H{first}<>"",H{first}<>"OK")'], fill=ALERT_FILL))
    ws.conditional_formatting.add("B4", FormulaRule(formula=['B4<>"OK"'], fill=ALERT_FILL))
    ws["J1"] = "Quote"
    ws["J1"].font = Font(bold=True)
    res = [
        ("Lines subtotal", f"=SUM({rg('F')})", USD),
        ("Setup fee", '=IF(K6=0,0,B3)', USD),
        ("Quote total", "=K2+K3", USD),
        ("Lines at the minimum", f'=COUNTIF({rg("G")},"Yes")', "0"),
        ("Lines priced", f"=COUNT({rg('F')})", "0"),
        ("Rows to check", f'=COUNTIF({rg("H")},"Check")', "0"),
        ("Recomputed: lines above minimum + minimum x lines at it",
         f'=SUMPRODUCT(--({rg("G")}="No"),{rg("E")},{rg("D")})*B1+B2*K5', USD),
        ("Check: subtotal = recomputed (half a cent per line)",
         '=IF(B4<>"OK","CHECK",IF(ABS(K2-K8)<=0.005*K6+0.000001,"OK","CHECK"))', "General"),
    ]
    for i, (label, f, fmt) in enumerate(res, start=2):
        ws[f"J{i}"], ws[f"K{i}"] = label, f
        ws[f"J{i}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws[f"K{i}"].number_format = fmt
        ws[f"K{i}"].fill = FORMULA_FILL
        ws[f"K{i}"].alignment = Alignment(horizontal="right", vertical="top")
    ws["J4"].font = Font(bold=True)
    ws.conditional_formatting.add("K9", FormulaRule(formula=['K9="CHECK"'], fill=ALERT_FILL))
    ws.freeze_panes = "A6"

    notes_sheet(wb, [
        "Tidy Tabs: Price a Custom Job by Square Foot With a Minimum Charge",
        "",
        "1. Quote tab: type your rate per square foot in B1, the minimum charge per line in B2 and a setup fee in B3 (yellow). B4 shows OK when the rate is above 0 and all three are numbers.",
        "2. Type each job with its width and height in inches and the quantity in A to D (yellow). Rows 6 to 25 have formulas ready.",
        "3. Area per piece E6: =ROUND(B6*C6/144,2). There are 144 square inches in a square foot, so 12 x 12 in is 1.00 sq ft and 24 x 36 in is 6.00 sq ft.",
        "4. Line price F6: =MAX($B$2,ROUND(E6*D6*$B$1,2)). MAX returns the minimum charge when area x quantity x rate comes to less.",
        "5. Minimum applied G6 shows Yes when the minimum raised the price. A line that comes to exactly the minimum shows No (the price is the same either way).",
        "6. Size check H6 shows Check when a width, height or quantity is blank, 0, negative, a fraction of a piece, or typed as text. Those rows get no price. Under 36 sq in: feet? warns that a size may have been typed in feet instead of inches (3 x 6 typed for a 3 x 6 ft banner is only 18 square inches); the row is still priced.",
        "7. K2 to K4 add the lines and the setup fee (the setup fee is added only when at least one line is priced). K9 recomputes the subtotal another way and shows OK when it matches within half a cent per line and B4 is OK.",
        "8. Arithmetic only. Sales tax, discounts and how to set your rate are not included.",
        "",
        "Shop, jobs and rates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb["Quote"].print_title_rows = "5:5"
    wb.save(OUT / "tidy-tabs-price-per-square-foot.xlsx")


def expiring_dates():
    wb = Workbook()
    ws = wb.active
    ws.title = "Items"
    header(ws, ["Item", "Type", "Expires", "Days left", "Status"], [30, 22, 13, 11, 11])
    names = [
        ("Oat milk lotion 8 oz", "Lotion"), ("Lavender hand lotion", "Lotion"),
        ("Soy wax flakes 10 lb", "Candle wax"), ("Beeswax pellets 5 lb", "Candle wax"),
        ("Fall market banner", "Booth banner order"), ("Holiday market banner", "Booth banner order"),
        ("Vendor ID card, Ana", "Market vendor ID card"), ("Vendor ID card, Ben", "Market vendor ID card"),
        ("Shea butter lotion", "Lotion"), ("Cedar candle wax 25 lb", "Candle wax"),
        ("Spring market banner", "Booth banner order"), ("Vendor ID card, Cleo", "Market vendor ID card"),
        ("Cucumber face lotion", "Lotion"), ("Paraffin blend wax", "Candle wax"),
        ("Pop-up shop banner", "Booth banner order"), ("Vendor ID card, Dev", "Market vendor ID card"),
        ("Unscented body lotion", "Lotion"), ("Pillar candle wax 5 lb", "Candle wax"),
        ("Gift fair banner", "Booth banner order"), ("Vendor ID card, Eli", "Market vendor ID card"),
    ]
    offsets = [-1, 0, 30, 31, 12, -30, 5, 90, 21, 29, 45, 8, 120, 14, 365, 2, 60, 33, 180, -12]
    asof = date(2026, 10, 19)
    last = 60
    for r in range(2, last + 1):
        if r - 2 < len(names):
            ws[f"A{r}"], ws[f"B{r}"] = names[r - 2]
            ws[f"C{r}"] = date.fromordinal(asof.toordinal() + offsets[r - 2])
        ws[f"D{r}"] = (f'=IF(OR(C{r}="",NOT(ISNUMBER(C{r})),NOT(ISNUMBER(Summary!$B$1))),"",'
                       f'C{r}-Summary!$B$1)')
        ws[f"E{r}"] = (f'=IF(C{r}="","",IF(OR(NOT(ISNUMBER(D{r})),NOT(ISNUMBER(Summary!$B$2))),"Check",'
                       f'IF(D{r}<0,"Expired",IF(D{r}<=Summary!$B$2,"Soon","OK"))))')
        ws[f"C{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = "0"
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    # Rules read the Status text, so they work without cross-sheet references.
    red = PatternFill("solid", bgColor="F8D7DA", fgColor="F8D7DA")
    amber = PatternFill("solid", bgColor="FFE699", fgColor="FFE699")
    ws.conditional_formatting.add(f"A2:E{last}", FormulaRule(formula=['$E2="Expired"'], fill=red, font=Font(bold=True, color="9C0006")))
    ws.conditional_formatting.add(f"A2:E{last}", FormulaRule(formula=['$E2="Soon"'], fill=amber))
    ws.conditional_formatting.add(f"E2:E{last}", FormulaRule(formula=['$E2="Check"'], fill=ALERT_FILL))

    sm = wb.create_sheet("Summary", 1)
    sm.column_dimensions["A"].width = 52
    sm.column_dimensions["B"].width = 14
    rng = f"Items!$C$2:$C${last}"
    st = f"Items!$E$2:$E${last}"
    res = [
        ("As-of date (type today's date)", date(2026, 10, 19), US_DATE),
        ("Warning window (days)", 30, "0"),
        ("Items expired (before the as-of date)", f'=COUNTIF({st},"Expired")', "0"),
        ("Items expiring soon (0 to the window)", f'=COUNTIF({st},"Soon")', "0"),
        ("Items OK (past the window)", f'=COUNTIF({st},"OK")', "0"),
        ("Rows to check (date typed as text)", f'=COUNTIF({st},"Check")', "0"),
        ("Rows with an Expires entry", f"=COUNTA({rng})", "0"),
        ("Dates from the as-of date through the window (COUNTIFS)",
         f'=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),COUNTIFS({rng},">="&B1,{rng},"<="&B1+B2),"")', "0"),
        ("Check: Expired + Soon + OK = real dates, COUNTIFS = Soon, nothing to check",
         f'=IF(NOT(AND(ISNUMBER(B1),ISNUMBER(B2))),"Check settings",IF(AND(B3+B4+B5=COUNT({rng}),B6=0,B4=B8),"OK","CHECK"))', "General"),
    ]
    for i, (label, v, fmt) in enumerate(res, start=1):
        sm[f"A{i}"], sm[f"B{i}"] = label, v
        sm[f"A{i}"].font = Font(bold=True)
        sm[f"B{i}"].number_format = fmt
        sm[f"B{i}"].alignment = Alignment(horizontal="right")
        sm[f"B{i}"].fill = INPUT_FILL if i <= 2 else FORMULA_FILL
    sm.conditional_formatting.add("B9", FormulaRule(formula=['B9<>"OK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Highlight Dates That Are Coming Up Soon",
        "",
        "1. Summary tab: type the as-of date in B1 (yellow) and how many days ahead to warn in B2. The sheet does not use TODAY, so type a new date when you open it.",
        "2. Items tab: one row per item with its Expires date in C (yellow). Rows 2 to 60 have formulas ready.",
        "3. Days left D2: =C2-Summary!$B$1. Negative means the date has passed.",
        "4. Status E2: Expired when Days left is below 0, Soon when it is 0 up to the warning window, OK after that. An item that expires on the as-of date is Soon; one that expires the day before is Expired.",
        "5. Colors: the rows are colored by conditional formatting rules that read the Status text (red for Expired, amber for Soon). The Status words also print in black and white.",
        "6. To build the rule yourself: select A2:E60, then Home > Conditional Formatting > New Rule > Use a formula, and type =$E2=\"Soon\". In Google Sheets use Format > Conditional formatting > Custom formula is. Microsoft and Google help pages describe the dialogs.",
        "7. A date typed as text (usually left-aligned) shows Check and is left out of the counts. A blank date has no status.",
        "8. Summary B8 counts dates from the as-of date through the window with COUNTIFS. B9 shows OK when Expired + Soon + OK equals the real dates, B8 equals the Soon count and no row needs checking.",
        "9. Dates only. This sheet does not know shelf life, safety or regulatory rules for any product.",
        "10. The print area covers the first 40 rows. Change the print area to print more.",
        "",
        "Items and dates are fictional. The sheet's formulas and the rule text were checked in LibreOffice Calc only; the colors were not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Items"].print_area = "A1:E41"
    wb.save(OUT / "tidy-tabs-expiring-dates.xlsx")


def validation_rules():
    wb = Workbook()
    ws = wb.active
    ws.title = "Entries"
    header(ws, ["Invoice #", "Date", "Qty"], [14, 14, 10])
    rows = [
        ("INV-1001", (10, 1), 4), ("INV-1002", (10, 2), 1), ("INV-1003", (10, 2), 12),
        ("INV-1004", (10, 5), 3), ("INV-1005", (10, 6), 6), ("INV-1006", (10, 7), 2),
        ("INV-1007", (10, 8), 8), ("INV-1008", (10, 9), 2.5), ("INV-1009", (10, 12), 5),
        ("INV-1006", (10, 13), 1), ("INV-1010", (10, 14), 20), ("INV-1011", (10, 15), 7),
    ]
    last = 200
    for r in range(2, last + 1):
        if r - 2 < len(rows):
            inv, (m, d), q = rows[r - 2]
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = inv, date(2026, m, d), q
        ws[f"B{r}"].number_format = US_DATE
        if r <= 60:
            for col in "ABC":
                ws[f"{col}{r}"].fill = INPUT_FILL
    rules = [
        ("A", f"COUNTIF($A$2:$A${last},A2)=1", "Duplicate invoice number",
         "This invoice number is already in the list. Type a new one.", "Invoice number", "Type a number not used yet, such as INV-1012."),
        ("B", "AND(ISNUMBER(B2),B2>=DATE(2026,1,1),B2<=DATE(2026,12,31))", "Not a 2026 date",
         "Type a real date from 01/01/2026 to 12/31/2026, such as 10/19/2026.", "Date", "MM/DD/YYYY, in 2026."),
        ("C", "IF(ISNUMBER(C2),AND(C2>0,C2=INT(C2)),FALSE)", "Not a whole quantity",
         "Type a whole number above 0, such as 3.", "Qty", "A whole number above 0."),
    ]
    for col, f, etitle, emsg, ptitle, pmsg in rules:
        dv = DataValidation(type="custom", formula1=f, allow_blank=True, errorStyle="stop",
                            showErrorMessage=True, showInputMessage=True,
                            errorTitle=etitle, error=emsg, promptTitle=ptitle, prompt=pmsg)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{last}")

    rc = wb.create_sheet("Rules check", 1)
    header(rc, ["Invoice #", "Invoice # unique?", "Date in 2026?", "Qty whole, above 0?", "Row result"],
           [14, 17, 14, 19, 12])
    for r in range(2, last + 1):
        empty = f'AND(Entries!A{r}="",Entries!B{r}="",Entries!C{r}="")'
        rc[f"A{r}"] = f'=IF({empty},"",IF(Entries!A{r}="","(blank)",Entries!A{r}))'
        rc[f"B{r}"] = f'=IF({empty},"",AND(Entries!A{r}<>"",COUNTIF(Entries!$A$2:$A${last},Entries!A{r})=1))'
        rc[f"C{r}"] = (f'=IF({empty},"",AND(ISNUMBER(Entries!B{r}),Entries!B{r}>=DATE(2026,1,1),'
                       f'Entries!B{r}<=DATE(2026,12,31)))')
        rc[f"D{r}"] = f'=IF({empty},"",IF(ISNUMBER(Entries!C{r}),AND(Entries!C{r}>0,Entries!C{r}=INT(Entries!C{r})),FALSE))'
        rc[f"E{r}"] = f'=IF({empty},"",IF(AND(B{r},C{r},D{r}),"OK","Fix"))'
        for col in "ABCDE":
            rc[f"{col}{r}"].fill = FORMULA_FILL
            rc[f"{col}{r}"].alignment = Alignment(horizontal="center")
    rc.conditional_formatting.add(f"A2:E{last}", FormulaRule(formula=['$E2="Fix"'], fill=ALERT_FILL))
    rc.column_dimensions["F"].width = 4
    rc.column_dimensions["G"].width = 46
    rc.column_dimensions["H"].width = 10
    rc["G1"] = "Counts and check"
    rc["G1"].font = Font(bold=True)
    res = [
        ("Rows with an entry", f"=SUMPRODUCT(--(E2:E{last}<>\"\"))"),
        ("Rows that pass all three rules", f'=COUNTIF(E2:E{last},"OK")'),
        ("Rows to fix", f'=SUMPRODUCT(--(E2:E{last}="Fix"))'),
        ("Rows failing the invoice rule", f"=COUNTIF(B2:B{last},FALSE)"),
        ("Rows failing the date rule", f"=COUNTIF(C2:C{last},FALSE)"),
        ("Rows failing the qty rule", f"=COUNTIF(D2:D{last},FALSE)"),
        ("Rows sharing an invoice number (second count)",
         f'=SUMPRODUCT((Entries!A2:A{last}<>"")*(COUNTIF(Entries!A2:A{last},Entries!A2:A{last})>1))'),
        ("Check: OK + Fix = rows, duplicate counts agree",
         '=IF(AND(H3+H4=H2,H8=SUMPRODUCT(--(B2:B' + str(last) + '=FALSE),--(Entries!A2:A' + str(last) + '<>""))),"OK","CHECK")'),
    ]
    for i, (label, f) in enumerate(res, start=2):
        rc[f"G{i}"], rc[f"H{i}"] = label, f
        rc[f"H{i}"].fill = FORMULA_FILL
        rc[f"H{i}"].alignment = Alignment(horizontal="right")
    rc.conditional_formatting.add("H9", FormulaRule(formula=['H9="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Block Duplicate Invoice Numbers and Bad Dates",
        "",
        "1. Entries tab: type Invoice #, Date and Qty in A to C (yellow). Rows 2 to 200 have the three validation rules.",
        "2. Select a column, then Data > Data Validation > Allow: Custom. The rules here are: A2 =COUNTIF($A$2:$A$200,A2)=1 (no repeats), B2 =AND(ISNUMBER(B2),B2>=DATE(2026,1,1),B2<=DATE(2026,12,31)) (a real 2026 date), C2 =IF(ISNUMBER(C2),AND(C2>0,C2=INT(C2)),FALSE) (a whole number above 0).",
        "3. Each rule is written for the first cell (A2, B2, C2). The cell references without $ shift down for every row. $A$2:$A$200 stays fixed.",
        "4. Each rule has a Stop error message and an input prompt, shown when you select the cell and when a bad value is typed.",
        "5. Validation only checks what you type. Pasting over a cell skips the rule, and values already in the cells are not checked. Rows 7 and 11 repeat INV-1006 and row 9 has a quantity of 2.5 on purpose, so you can see them on the Rules check tab.",
        "6. Rules check tab: the same three tests as TRUE or FALSE for every row, a Row result (OK or Fix) and counts. Rows to fix (H4) is the number to clean up after a paste. Both rows of a duplicated invoice number fail the invoice test.",
        "7. H9 shows OK when OK rows plus Fix rows equal the rows with an entry and two counts of duplicate rows agree.",
        "8. COUNTIF ignores upper and lower case (inv-1006 repeats INV-1006) and does not ignore spaces. The date rule allows 2026 only and a date with a time of day on 12/31/2026 fails; edit the DATE(...) parts for other years.",
        "9. To change a rule, select the column, open Data Validation again and edit the formula.",
        "",
        "Invoice numbers, dates and quantities are fictional. The rule formulas and the Rules check tab were checked in LibreOffice Calc only; the blocking itself, typing a bad value, was not tried in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx. In Sheets use Data > Data validation and the Custom formula is option.",
    ])
    finish_letter(wb)
    wb["Entries"].print_area = "A1:C41"
    wb["Rules check"].print_area = "A1:H41"
    wb.save(OUT / "tidy-tabs-validation-rules.xlsx")


def sum_with_or():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Date", "Order ID", "Channel", "Amount"], [12, 11, 17, 12])
    ch = ["Etsy", "Shopify", "Farmers market", "Wholesale"]
    rows = [
        ((10, 1), "Etsy", 18.00), ((10, 1), "Shopify", 42.50), ((10, 2), "Farmers market", 26.00),
        ((10, 2), "Etsy", 64.00), ((10, 3), "Wholesale", 240.00), ((10, 3), "Shopify", 31.00),
        ((10, 4), "Farmers market", 55.00), ((10, 5), "Etsy", 22.50), ((10, 5), "Shopify", 75.00),
        ((10, 6), "Farmers market", 18.00), ((10, 7), "Wholesale", 180.00), ((10, 7), "Etsy", 49.00),
        ((10, 8), "Shopify", 28.00), ((10, 9), "Farmers market", 62.00), ((10, 9), "Etsy", 36.00),
        ((10, 10), "Shopify", 90.00), ((10, 11), "Wholesale", 120.00), ((10, 12), "Etsy", 15.50),
        ((10, 12), "Farmers market", 33.00), ((10, 13), "Shopify", 24.00), ((10, 14), "Etsy", 52.00),
        ((10, 14), "Wholesale", 95.00), ((10, 15), "Shopify", 58.00), ((10, 15), "Farmers market", 41.00),
    ]
    last = 200
    for r in range(2, last + 1):
        if r - 2 < len(rows):
            (m, d), c, a = rows[r - 2]
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = date(2026, m, d), f"#{3101 + r - 2}", c, a
        ws[f"A{r}"].number_format = US_DATE
        ws[f"D{r}"].number_format = USD
        if r <= 60:
            for col in "ABCD":
                ws[f"{col}{r}"].fill = INPUT_FILL
    dvc = DataValidation(type="list", formula1='"' + ",".join(ch) + '"', allow_blank=True)
    ws.add_data_validation(dvc)
    dvc.add(f"C2:C{last}")

    t = wb.create_sheet("Totals", 1)
    t.column_dimensions["A"].width = 58
    t.column_dimensions["B"].width = 16
    C, D = f"Orders!$C$2:$C${last}", f"Orders!$D$2:$D${last}"
    pick = '(' + C + '=$B$1)*($B$1<>"")+(' + C + '=$B$2)*($B$2<>"")>0'
    res = [
        ("A1", "Channel 1", "B1", "Etsy", None),
        ("A2", "Channel 2", "B2", "Shopify", None),
        ("A3", "Large order threshold (amount over)", "B3", 50, USD),
        ("A5", "Total: SUMIFS for channel 1 + SUMIFS for channel 2", "B5",
         f"=SUMIFS({D},{C},B1)+SUMIFS({D},{C},B2)", USD),
        ("A6", 'Total: SUM(SUMIFS(...,{"Etsy","Shopify"})), names typed in the formula', "B6",
         f'=SUM(SUMIFS({D},{C},{{"Etsy","Shopify"}}))', USD),
        ("A7", "Total: SUMPRODUCT, a row counts once", "B7", f"=SUMPRODUCT(--({pick}),{D})", USD),
        ("A8", "Orders: COUNTIF + COUNTIF", "B8", f"=COUNTIF({C},B1)+COUNTIF({C},B2)", "0"),
        ("A9", "Orders: SUMPRODUCT, a row counts once", "B9", f"=SUMPRODUCT(--({pick}))", "0"),
        ("A10", "All other channels (not channel 1 or 2)", "B10",
         f'=SUMIFS({D},{C},"<>"&B1,{C},"<>"&B2)', USD),
        ("A11", "Everything minus the other channels", "B11", f"=SUM({D})-B10", USD),
        ("A12", "Either channel AND amount over the threshold", "B12",
         f"=SUMPRODUCT(--({pick}),--({D}>B3),{D})", USD),
        ("A13", "Check: the four totals agree",
         "B13", '=IF(OR(B1="",B2=""),"Pick two channels",IF(B1=B2,"Same channel twice",'
         'IF(AND(ABS(B5-B6)<0.005,ABS(B7-B6)<0.005,ABS(B11-B6)<0.005,B8=B9),"OK","CHECK")))', "General"),
    ]
    for la, label, va, v, fmt in res:
        t[la], t[va] = label, v
        t[la].font = Font(bold=True)
        t[va].alignment = Alignment(horizontal="right")
        t[va].fill = INPUT_FILL if va in ("B1", "B2", "B3") else FORMULA_FILL
        if fmt:
            t[va].number_format = fmt
    for va in ("B1", "B2"):
        dv = DataValidation(type="list", formula1='"' + ",".join(ch) + '"', allow_blank=True)
        t.add_data_validation(dv)
        dv.add(va)
    t.conditional_formatting.add("B13", FormulaRule(formula=['B13<>"OK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Sum Sales From Either of Two Channels (OR Logic)",
        "",
        "1. Orders tab: Date, Order ID, Channel and Amount in A to D (yellow). Formulas read rows 2 to 200.",
        "2. Totals tab: pick the two channels in B1 and B2 (yellow dropdowns). B3 is the amount used in B12.",
        "3. SUMIFS cannot take two values for one column, so B5 adds two SUMIFS: =SUMIFS(Orders!D2:D200,Orders!C2:C200,B1)+SUMIFS(Orders!D2:D200,Orders!C2:C200,B2).",
        "4. B6 sends both names in one array constant: =SUM(SUMIFS(Orders!D2:D200,Orders!C2:C200,{\"Etsy\",\"Shopify\"})). The names are typed inside the formula, so edit them there if you change B1 or B2.",
        "5. B7 uses SUMPRODUCT: =SUMPRODUCT(--((Orders!C2:C200=B1)*(B1<>\"\")+(Orders!C2:C200=B2)*(B2<>\"\")>0),Orders!D2:D200). The --( ) turns TRUE and FALSE into 1 and 0. Each row counts once. The (B1<>\"\") parts stop empty channel cells from matching the empty rows below your data.",
        "6. B8 and B9 count the orders the same two ways. B10 totals the other channels, so B11 (everything minus the others) should equal B5 to B7.",
        "7. B12 adds an AND to the OR: either channel, and an amount over B3.",
        "8. If channel 1 and channel 2 are the same, B5 adds that channel twice, while B7 and B6's names do not. B13 says Same channel twice.",
        "9. Matching ignores upper and lower case but not spaces: \"Etsy \" with a trailing space matches nothing. An amount typed as text is skipped by all methods.",
        "10. Sums and counts only. The print area on the Orders tab covers the first 40 rows.",
        "",
        "Shop, orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Orders"].print_area = "A1:D41"
    wb.save(OUT / "tidy-tabs-sum-with-or.xlsx")


def fill_blanks_above():
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    header(ws, ["Vendor", "Product", "Amount"], [24, 26, 12])
    groups = [
        ("Pine Hollow Candles", [("Cedar candle", 54.00), ("Balsam candle", 48.00), ("Wick refill pack", 18.50)]),
        ("Salt Marsh Soap", [("Sea salt bar", 30.00), ("Oat bar", 27.50), ("Gift set", 66.00), ("Soap dish", 21.00)]),
        ("Willow & Wick", [("Birch candle", 40.00), ("Match box", 12.00)]),
        ("Bramble Jam Co.", [("Blackberry jam", 36.00), ("Peach jam", 33.00), ("Jam sampler", 45.00)]),
        ("Sunny Creek Pottery", [("Mug", 64.00), ("Bowl", 52.00), ("Vase", 85.00)]),
    ]
    flat = []
    for v, items in groups:
        for i, (p, a) in enumerate(items):
            flat.append((v if i == 0 else None, p, a))
    last = 60
    for r in range(2, last + 1):
        if r - 2 < len(flat):
            ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = flat[r - 2]
        ws[f"C{r}"].number_format = USD
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL

    fl = wb.create_sheet("Filled", 1)
    header(fl, ["Vendor (fill down)", "Vendor (fill down, spaces too)", "Product", "Amount"], [22, 24, 24, 12])
    for r in range(2, last + 1):
        empty = f'AND(Report!A{r}="",Report!B{r}="",Report!C{r}="")'
        if r == 2:
            fl[f"A{r}"] = f'=IF({empty},"",IF(Report!A{r}="","",Report!A{r}))'
            fl[f"B{r}"] = f'=IF({empty},"",TRIM(Report!A{r}))'
        else:
            fl[f"A{r}"] = f'=IF({empty},"",IF(Report!A{r}="",A{r - 1},Report!A{r}))'
            fl[f"B{r}"] = f'=IF({empty},"",IF(TRIM(Report!A{r})="",B{r - 1},TRIM(Report!A{r})))'
        fl[f"C{r}"] = f'=IF(Report!B{r}="","",Report!B{r})'
        fl[f"D{r}"] = f'=IF(Report!C{r}="","",Report!C{r})'
        fl[f"D{r}"].number_format = USD
        for col in "ABCD":
            fl[f"{col}{r}"].fill = FORMULA_FILL
    fl.column_dimensions["E"].width = 4
    for col, w in (("F", 24), ("G", 16), ("H", 16), ("I", 16)):
        fl.column_dimensions[col].width = w
    for col, n in zip("FGHI", ["Vendor (type names)", "Total on Report", "Total, fill down", "Total, spaces too"]):
        c = fl[f"{col}1"]
        c.value, c.fill, c.font = n, HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    vendors = [g[0] for g in groups]
    for r in range(2, 9):
        fl[f"F{r}"] = vendors[r - 2] if r - 2 < len(vendors) else None
        fl[f"F{r}"].fill = INPUT_FILL
        fl[f"G{r}"] = f'=IF(F{r}="","",SUMIFS(Report!$C$2:$C${last},Report!$A$2:$A${last},F{r}))'
        fl[f"H{r}"] = f'=IF(F{r}="","",SUMIFS($D$2:$D${last},$A$2:$A${last},F{r}))'
        fl[f"I{r}"] = f'=IF(F{r}="","",SUMIFS($D$2:$D${last},$B$2:$B${last},F{r}))'
        for col in "GHI":
            fl[f"{col}{r}"].fill = FORMULA_FILL
            fl[f"{col}{r}"].number_format = USD
    used = f'(((Report!$B$2:$B${last}<>"")+(Report!$C$2:$C${last}<>""))>0)'
    chk = [
        ("Rows with a product or amount", f"=SUMPRODUCT(--{used})", "0"),
        ("Rows still blank in the vendor column (fill down)", f'=SUMPRODUCT(--{used},--($A$2:$A${last}=""))', "0"),
        ("Rows still blank (spaces too)", f'=SUMPRODUCT(--{used},--($B$2:$B${last}=""))', "0"),
        ("Total on Report", f"=SUM(Report!C2:C{last})", USD),
        ("Total on Filled", f"=SUM(D2:D{last})", USD),
        ("Check: no blanks left, totals equal",
         '=IF(AND(I11=0,I12=0,ABS(I13-I14)<0.005),"OK","CHECK")', "General"),
    ]
    for i, (label, f, fmt) in enumerate(chk, start=10):
        fl[f"F{i}"], fl[f"I{i}"] = label, f
        fl[f"F{i}"].font = Font(bold=True)
        fl[f"I{i}"].number_format = fmt
        fl[f"I{i}"].fill = FORMULA_FILL
        fl[f"I{i}"].alignment = Alignment(horizontal="right")
    fl["F9"] = "Checks"
    fl["F9"].font = Font(bold=True, size=12)
    fl.conditional_formatting.add("I15", FormulaRule(formula=['I15="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Fill Blank Cells With the Value Above",
        "",
        "1. Report tab: paste or type the report in A to C (yellow). Vendor is typed on the first row of each group only, like many exports.",
        "2. Filled tab, column A: =IF(Report!A3=\"\",A2,Report!A3) takes the vendor from the row above when the Report cell is blank. The first row is =IF(Report!A2=\"\",\"\",Report!A2) because there is no row above it (a bare =Report!A2 would show 0 for an empty cell).",
        "3. Column B does the same but treats a cell that holds only a space as blank, and trims stray spaces: =IF(TRIM(Report!A3)=\"\",B2,TRIM(Report!A3)).",
        "4. Both columns return blank for a row where the vendor, product and amount are all empty. A fully blank row ends a group: the next row needs its own vendor name.",
        "5. A blank in the very first data row stays blank. Fix it on the Report tab.",
        "6. Columns F to I: type vendor names in F (yellow). G is the total on the Report tab (the blank-labeled rows count for nobody), H and I are the totals from the filled columns.",
        "7. I15 shows OK when no used row is blank in either filled column and the amount totals are equal.",
        "8. To keep the result and drop the formulas: select Filled!A2:D60, Copy, then Paste Special > Values only.",
        "9. Menu route in Excel, no formulas: select the range, Home > Find & Select > Go To Special > Blanks, type = and the cell above, then press Ctrl+Enter. The steps are from Microsoft help and were not run by us.",
        "10. The Filled tab reads Report rows 2 to 60. Copy the formula row down for more.",
        "",
        "Vendors, products and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Report"].print_area = "A1:C41"
    wb["Filled"].print_area = "A1:I41"
    wb.save(OUT / "tidy-tabs-fill-blanks-above.xlsx")


def mailing_labels():
    wb = Workbook()
    ad = wb.active
    ad.title = "Addresses"
    header(ad, ["Name", "Street", "City", "State", "ZIP"], [22, 26, 18, 7, 9])
    people = [
        ("Maya Alvarez", "14 Beacon Row", "Boston", "MA", "02108"),
        ("Tom Brennan", "220 Willow Ave, Apt 3", "Hoboken", "NJ", "07030"),
        ("Priya Shah", "9 Orchard Lane", "Portland", "ME", "04101"),
        ("Jordan Lee", "1800 Maple Street", "Austin", "TX", "78701"),
        ("Dana Whitfield", "75 Harbor View Road", "Providence", "RI", "02903"),
        ("Carlos Ortega", "412 Elm Court", "Tucson", "AZ", "85701"),
        ("Ellen Park", "33 Mill Street", "Burlington", "VT", "05401"),
        ("Sam Okafor", "6 Cedar Hollow Drive", "Madison", "WI", "53703"),
        ("Rosa Delgado", "501 Pine Street", "Seattle", "WA", "98101"),
        ("Ben Kowalski", "88 River Road", "Pittsburgh", "PA", "15222"),
        ("Hannah Cho", "27 Lakeshore Boulevard, Suite 12", "Chicago", "IL", "60601"),
        ("Leo Martin", "3 Fairview Terrace", "Newark", "NJ", "07102"),
        ("Grace Nguyen", "640 Oak Avenue", "Denver", "CO", "80202"),
        ("Miles Foster", "19 Quarry Lane", "Concord", "NH", "03301"),
        ("Aisha Rahman", "250 Summit Street", "Hartford", "CT", "06103"),
        ("Pete Sullivan", "71 Bay Street", "Savannah", "GA", "31401"),
        ("Nora Jensen", "12 Prairie Way", "Omaha", "NE", "68102"),
        ("Victor Reyes", "905 Canal Street", "New Orleans", "LA", "70130"),
    ]
    for r in range(2, 41):
        if r - 2 < len(people):
            for col, v in zip("ABCDE", people[r - 2]):
                ad[f"{col}{r}"] = v
        ad[f"E{r}"].number_format = "@"
        for col in "ABCDE":
            ad[f"{col}{r}"].fill = INPUT_FILL

    lb = wb.create_sheet("Labels", 1)
    from openpyxl.styles import Border, Side
    dot = Side(style="hair", color="999999")
    for col in "ABC":
        lb.column_dimensions[col].width = 32
    for r in range(1, 7):
        lb.row_dimensions[r].height = 100
        for c in range(1, 4):
            n = f"((ROW()-1)*3+COLUMN())"
            parts = [f"INDEX(Addresses!${x}$2:${x}$19,{n})" for x in "ABCDE"]
            cell = lb.cell(row=r, column=c)
            cell.value = (f'=IFERROR(IF({parts[0]}&{parts[1]}&{parts[2]}="","",'
                          f'{parts[0]}&CHAR(10)&{parts[1]}&CHAR(10)&{parts[2]}&", "&{parts[3]}&" "&{parts[4]}),"")')
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left", indent=1)
            cell.border = Border(left=dot, right=dot, top=dot, bottom=dot)
    lb.column_dimensions["D"].width = 3
    lb.column_dimensions["E"].width = 44
    lb.column_dimensions["F"].width = 10
    res = [
        ("Addresses listed (rows 2 to 19)", '=SUMPRODUCT(--(Addresses!A2:A19&Addresses!B2:B19&Addresses!C2:C19<>""))', "0"),
        ("Label cells filled", '=COUNTIF(A1:C6,"?*")', "0"),
        ("Addresses past row 19 (not on this sheet)", '=SUMPRODUCT(--(Addresses!A20:A200&Addresses!B20:B200&Addresses!C20:C200<>""))', "0"),
        ("Check: every address has a label", '=IF(AND(F2=F3,F4=0),"OK","CHECK")', "General"),
    ]
    lb["E1"] = "Counts and check (not printed)"
    lb["E1"].font = Font(bold=True)
    for i, (label, f, fmt) in enumerate(res, start=2):
        lb[f"E{i}"], lb[f"F{i}"] = label, f
        lb[f"E{i}"].alignment = Alignment(wrap_text=False, vertical="center")
        lb[f"F{i}"].alignment = Alignment(horizontal="right", vertical="center")
        lb[f"F{i}"].fill = FORMULA_FILL
    lb.conditional_formatting.add("F5", FormulaRule(formula=['F5="CHECK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Mailing Labels From an Address List",
        "",
        "1. Addresses tab: Name, Street, City, State and ZIP in A to E (yellow), one person per row. The ZIP column is formatted as text so 02108 keeps its zero. If you paste ZIPs that already lost the zero, retype them.",
        "2. Labels tab: 3 labels across by 6 down on one US Letter page (18 labels), margins 0.5 in. Label 1 is the top left, then across, then the next row.",
        "3. Each label cell holds the same formula, copied across the grid. A1 picks address number (ROW()-1)*3+COLUMN() with INDEX: =INDEX(Addresses!$A$2:$A$19,(ROW()-1)*3+COLUMN()) and joins Name, Street and City, State ZIP with CHAR(10) (a line break). Wrap text is on.",
        "4. When the list has fewer than 18 addresses, the cells past the end are blank. A blank row on the Addresses tab leaves its label blank.",
        "5. Only rows 2 to 19 of the Addresses tab are used. With more than 18 addresses the extra ones are not shown: Labels!F4 counts them and F5 says CHECK. Print 18, then move the next 18 up to rows 2 to 19.",
        "6. Print on plain paper first. Hold it over the label sheet against a window to see if the grid lines up. Sheets from different brands have different layouts and this grid is not an Avery or other brand template. Change column widths and row heights on the Labels tab to adjust.",
        "7. Labels!F5 shows OK when the label cells filled equal the addresses listed and none are past row 19. The counts are outside the print area (A1:C6).",
        "8. Layout only. For addressing rules, see the USPS help pages.",
        "",
        "Names and addresses are fictional. Checked in LibreOffice Calc only (formulas and PDF export); not opened in Excel or Google Sheets, and not printed on label paper.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    lb.print_title_rows = None
    lb.print_area = "A1:C6"
    lb.page_margins.left = lb.page_margins.right = lb.page_margins.top = lb.page_margins.bottom = 0.5
    lb.print_options.horizontalCentered = True
    wb["Addresses"].print_area = "A1:E41"
    wb.save(OUT / "tidy-tabs-mailing-labels.xlsx")


def months_to_pay_off():
    wb = Workbook()
    ws = wb.active
    ws.title = "Payoff"
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 60
    rows = [
        (1, "Balance", 5000, USD, True, "Typed. A kiln or laptop balance (fictional)."),
        (2, "Yearly interest rate", 0.08, "0.00%", True, "Type 0% if there is no interest."),
        (3, "Monthly payment", 400, USD, True, ""),
        (4, "First payment date", date(2026, 11, 1), US_DATE, True, "Typed. The sheet does not use TODAY."),
        (5, "Months to pay off",
         '=IF(OR(NOT(ISNUMBER(B1)),NOT(ISNUMBER(B2)),NOT(ISNUMBER(B3)),NOT(ISNUMBER(B4))),"Check inputs",'
         'IF(OR(B1<=0,B2<0,B3<=0),"Check inputs",IF(B3<=B1*B2/12,"Payment too low",'
         'ROUNDUP(NPER(B2/12,-B3,B1),0))))', "0", False, "Formula: =ROUNDUP(NPER(B2/12,-B3,B1),0), plus guards"),
        (6, "Last payment date", '=IF(ISNUMBER(B5),EDATE(B4,B5-1),"")', US_DATE, False, ""),
        (7, "Last payment", '=IF(ISNUMBER(B5),IF(B5<=60,INDEX(Schedule!D2:D61,B5),"Over 60 months"),"")', USD, False,
         "The part month: smaller than the regular payment."),
        (8, "Total paid", '=IF(ISNUMBER(B7),B3*(B5-1)+B7,"")', USD, False, ""),
        (9, "Interest paid", '=IF(ISNUMBER(B8),ROUND(B8-B1,2),"")', USD, False, ""),
        (10, "Check: schedule agrees with NPER",
         '=IF(NOT(ISNUMBER(B5)),"No result",IF(B5>60,"Over 60 months",'
         'IF(AND(COUNT(Schedule!D2:D61)=B5,INDEX(Schedule!E2:E61,B5)=0),"OK","CHECK")))', "General", False,
         "OK when the Schedule tab has exactly that many payments and ends at $0.00."),
        (12, "Save for a goal", None, "General", False, ""),
        (13, "Goal amount", 3000, USD, True, "A booth deposit (fictional)."),
        (14, "Monthly deposit", 250, USD, True, ""),
        (15, "Yearly rate on savings", 0.02, "0.00%", True, "Type 0% for none."),
        (16, "Months to reach the goal",
         '=IF(OR(NOT(ISNUMBER(B13)),NOT(ISNUMBER(B14)),NOT(ISNUMBER(B15))),"Check inputs",'
         'IF(OR(B13<=0,B14<=0,B15<0),"Check inputs",ROUNDUP(NPER(B15/12,-B14,0,B13),0)))', "0", False,
         "Formula: =ROUNDUP(NPER(B15/12,-B14,0,B13),0)"),
        (17, "Check: FV reaches the goal in that month, not before",
         '=IF(ISNUMBER(B16),IF(AND(FV(B15/12,B16,-B14)>=B13-0.005,FV(B15/12,B16-1,-B14)<B13),"OK","CHECK"),"")',
         "General", False, "Deposits at the end of each month."),
    ]
    for r, label, v, fmt, inp, note in rows:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = Font(bold=True, size=12 if r == 12 else 11)
        if v is not None:
            ws[f"B{r}"] = v
            ws[f"B{r}"].number_format = fmt
            ws[f"B{r}"].fill = INPUT_FILL if inp else FORMULA_FILL
            ws[f"B{r}"].alignment = Alignment(horizontal="right")
        if note:
            ws[f"C{r}"] = note
    for r in (10, 17):
        ws.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f'AND(B{r}<>"OK",B{r}<>"")'], fill=ALERT_FILL))
    for r in (5, 16):
        ws.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f'NOT(ISNUMBER(B{r}))'], fill=ALERT_FILL))

    sc = wb.create_sheet("Schedule", 1)
    header(sc, ["Month", "Starting balance", "Interest", "Payment", "Ending balance", "Payment date"],
           [8, 16, 12, 12, 16, 14])
    ok = "ISNUMBER(Payoff!$B$5)"
    for n in range(1, 61):
        r = n + 1
        sc[f"A{r}"] = n
        if n == 1:
            sc[f"B{r}"] = f'=IF(NOT({ok}),"",Payoff!$B$1)'
        else:
            sc[f"B{r}"] = f'=IF(E{r - 1}="","",IF(E{r - 1}<=0,"",E{r - 1}))'
        sc[f"C{r}"] = f'=IF(B{r}="","",ROUND(B{r}*Payoff!$B$2/12,2))'
        sc[f"D{r}"] = f'=IF(B{r}="","",MIN(Payoff!$B$3,ROUND(B{r}+C{r},2)))'
        sc[f"E{r}"] = f'=IF(B{r}="","",ROUND(B{r}+C{r}-D{r},2))'
        sc[f"F{r}"] = f'=IF(B{r}="","",EDATE(Payoff!$B$4,A{r}-1))'
        for col in "BCDE":
            sc[f"{col}{r}"].number_format = USD
        sc[f"F{r}"].number_format = US_DATE
        for col in "ABCDEF":
            sc[f"{col}{r}"].fill = FORMULA_FILL

    notes_sheet(wb, [
        "Tidy Tabs: How Many Months to Pay Off a Balance",
        "",
        "1. Payoff tab: type the balance (B1), the yearly interest rate (B2), the monthly payment (B3) and the first payment date (B4). They are yellow.",
        "2. B5 is =ROUNDUP(NPER(B2/12,-B3,B1),0). NPER wants the monthly rate (yearly rate divided by 12), the payment as a negative number, and the balance as the present value. ROUNDUP is there because a part month still needs a payment.",
        "3. If the payment is not bigger than one month of interest (B1*B2/12), the balance never goes down and NPER returns #NUM!. B5 shows Payment too low instead. A blank or text input, a rate below 0, or a balance or payment of 0 or less shows Check inputs.",
        "4. B6 is the last payment date, =EDATE(B4,B5-1). B7 is the last payment, taken from the Schedule tab. B8 is B3*(B5-1)+B7. B9 is B8-B1.",
        "5. Schedule tab: 60 rows. Each month: interest = balance x yearly rate / 12 rounded to cents, payment = the smaller of your payment and balance plus interest, ending balance = balance + interest - payment. B10 on the Payoff tab shows OK when the schedule has exactly the NPER month count and ends at $0.00.",
        "6. Longer than 60 months: B5 still works, but the Schedule tab stops at row 61 and B7 to B10 say Over 60 months. Copy the last Schedule row down to extend it.",
        "7. Save for a goal (B13 to B17): =ROUNDUP(NPER(B15/12,-B14,0,B13),0) is the number of monthly deposits to reach the goal, with the first deposit at the end of month 1. B17 checks it with FV: the goal is reached in that month and not the month before.",
        "8. The schedule rounds each month's interest to cents, so a real lender's statement can differ by cents or a month in rare cases. Arithmetic only: this is not borrowing advice and it does not follow any lender's APR rules.",
        "",
        "The balance, rates and dates are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Payoff"].print_title_rows = None
    wb["Schedule"].print_area = "A1:F61"
    wb.save(OUT / "tidy-tabs-months-to-pay-off.xlsx")


def sum_across_monthly_sheets():
    wb = Workbook()
    months = {
        "Jan": (3150.00, 820.40, 252.00),
        "Feb": (2875.50, 640.25, 230.04),
        "Mar": (3980.00, 1105.80, 318.40),
        "Apr": (4410.75, 960.10, 352.86),
    }
    first = True
    for name, vals in months.items():
        ws = wb.active if first else wb.create_sheet(name)
        ws.title = name
        first = False
        header(ws, ["Line", "Amount"], [16, 14])
        for i, (label, v) in enumerate(zip(["Sales", "Supplies", "Fees"], vals), start=2):
            ws[f"A{i}"], ws[f"B{i}"] = label, v
            ws[f"B{i}"].number_format = USD
            ws[f"B{i}"].fill = INPUT_FILL

    sm = wb.create_sheet("Summary")
    header(sm, ["Line", "Plus version", "3D SUM (Excel)", "INDIRECT version"], [50, 18, 18, 18])
    for i, label in enumerate(["Sales", "Supplies", "Fees"], start=2):
        sm[f"A{i}"] = label
        sm[f"B{i}"] = f"=Jan!B{i}+Feb!B{i}+Mar!B{i}+Apr!B{i}"
        sm[f"C{i}"] = f"=SUM(Jan:Apr!B{i})"
        sm[f"D{i}"] = f"={'BCD'[i - 2]}21"
        for col in "BCD":
            sm[f"{col}{i}"].number_format = USD
            sm[f"{col}{i}"].fill = FORMULA_FILL
    sm["A6"] = "The plus version has no 3D reference, so use it in Google Sheets. 3D SUM is for Excel and LibreOffice."
    sm["A6"].font = Font(italic=True)
    sm["A7"] = "Months by tab name (INDIRECT)"
    sm["A7"].font = Font(bold=True, size=12)
    for col, t in zip("ABCD", ["Tab name (type it)", "Sales", "Supplies", "Fees"]):
        c = sm[f"{col}8"]
        c.value, c.fill, c.font = t, HEADER_FILL, HEADER_FONT
    for r in range(9, 21):
        sm[f"A{r}"].fill = INPUT_FILL
        if r - 9 < 4:
            sm[f"A{r}"] = list(months)[r - 9]
        for col, src in zip("BCD", (2, 3, 4)):
            ref = f'INDIRECT("\'"&$A{r}&"\'!B{src}")'
            sm[f"{col}{r}"] = f'=IF($A{r}="","",IF(ISERROR({ref}),"Check sheet name",{ref}))'
            sm[f"{col}{r}"].number_format = USD
            sm[f"{col}{r}"].fill = FORMULA_FILL
    sm["A21"] = "Total"
    sm["A21"].font = Font(bold=True)
    for col in "BCD":
        sm[f"{col}21"] = f"=SUM({col}9:{col}20)"
        sm[f"{col}21"].number_format = USD
        sm[f"{col}21"].font = Font(bold=True)
        sm[f"{col}21"].fill = FORMULA_FILL
    sm["A23"] = "Check: plus version equals INDIRECT totals"
    sm["B23"] = ('=IFERROR(IF(AND(COUNT(B9:D20)=3*COUNTA(A9:A20),ABS(B2-D2)<0.005,'
                 'ABS(B3-D3)<0.005,ABS(B4-D4)<0.005),"OK","CHECK"),"CHECK")')
    sm["A24"] = "Check: 3D SUM equals plus version (Excel, LibreOffice)"
    sm["B24"] = ('=IF(ISERROR(C2+C3+C4),"Not supported here",IFERROR(IF(AND(ABS(B2-C2)<0.005,'
                 'ABS(B3-C3)<0.005,ABS(B4-C4)<0.005),"OK","CHECK"),"CHECK"))')
    for r in (23, 24):
        sm[f"A{r}"].font = Font(bold=True)
        sm[f"B{r}"].fill = FORMULA_FILL
        sm[f"B{r}"].alignment = Alignment(horizontal="right")
    sm.conditional_formatting.add("B23", FormulaRule(formula=['B23<>"OK"'], fill=ALERT_FILL))
    sm.conditional_formatting.add("B24", FormulaRule(formula=['AND(B24<>"OK",B24<>"Not supported here")'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Add the Same Cell Across Monthly Sheets",
        "",
        "1. Jan, Feb, Mar and Apr tabs have the same layout: Sales in B2, Supplies in B3, Fees in B4 (yellow). Keep every month in the same cells.",
        "2. Summary B2: =Jan!B2+Feb!B2+Mar!B2+Apr!B2. This plain version works in Excel, Google Sheets and LibreOffice. To add a month, add + and the new tab's cell. A text amount such as n/a makes + show #VALUE!. Text that looks like a number, such as $100 typed as text, is converted by + in LibreOffice but SUM skips it, so the plus and 3D totals then differ and B24 shows CHECK.",
        "3. Summary C2: =SUM(Jan:Apr!B2). This is a 3D reference: it adds B2 on every tab from Jan to Apr, by tab position. In Excel a tab dragged between Jan and Apr is included and a tab outside them is not. SUM ignores text amounts. Google Sheets does not support 3D references, so column C shows an error there (see Google's help for Sheets). Do not rely on it if you use Sheets.",
        "4. Summary A9:A20 and B9:D21: type each tab name in column A (yellow). B9 is =IF($A9=\"\",\"\",IF(ISERROR(INDIRECT(\"'\"&$A9&\"'!B2\")),\"Check sheet name\",INDIRECT(\"'\"&$A9&\"'!B2\"))). INDIRECT builds the cell address from text and, unlike a 3D reference, exists in Excel, Google Sheets and LibreOffice (we ran it in LibreOffice only). A misspelled tab name shows Check sheet name instead of #REF!.",
        "5. INDIRECT reads the typed name. If you rename a tab, the plus and 3D formulas follow the new name in Excel, but the INDIRECT table keeps the old text and shows Check sheet name until you retype the name in column A.",
        "6. Summary B23 shows OK when the plus totals equal the INDIRECT totals and every typed tab name returned numbers. B24 shows OK when the 3D totals equal the plus totals, or Not supported here when the app has no 3D reference.",
        "7. Sums only. Sales, supplies and fees here are sample amounts, not bookkeeping or tax advice.",
        "",
        "The shop and the amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx. Column C will not calculate in Sheets.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-sum-across-monthly-sheets.xlsx")


def outlier_order_amounts():
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order ID", "Date", "Amount", "Flag"], [12, 13, 12, 18])
    amounts = [42.50, 58.00, 49.99, 61.25, 37.00, 54.75, 68.00, 45.50, 52.00, 39.95,
               500.00, 71.50, 48.00, 56.25, 63.00, 44.00, 50.00, 35.50, 59.90, 66.75,
               4.90, 47.25, 53.50, 40.00, 57.00, 62.50, 46.00, 74.00, 51.25, 38.75]
    last = 100
    for r in range(2, last + 1):
        i = r - 2
        if i < len(amounts):
            ws[f"A{r}"] = f"ORD-{4101 + i}"
            ws[f"B{r}"] = date.fromordinal(date(2026, 9, 14).toordinal() + i)
            ws[f"C{r}"] = amounts[i]
        ws[f"B{r}"].number_format = US_DATE
        ws[f"C{r}"].number_format = USD
        ws[f"D{r}"] = (f'=IF(C{r}="","",IF(NOT(ISNUMBER(C{r})),"Check",IF(NOT(AND(ISNUMBER(Fences!$B$6),ISNUMBER(Fences!$B$7))),"Check",'
                       f'IF(OR(C{r}>Fences!$B$6,C{r}<Fences!$B$7),"Review","OK"))))')
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        ws[f"D{r}"].fill = FORMULA_FILL
    amber = PatternFill("solid", bgColor="FFE699", fgColor="FFE699")
    ws.conditional_formatting.add(f"A2:D{last}", FormulaRule(formula=['$D2="Review"'], fill=amber, font=Font(bold=True)))
    ws.conditional_formatting.add(f"D2:D{last}", FormulaRule(formula=['$D2="Check"'], fill=ALERT_FILL))

    fn = wb.create_sheet("Fences", 1)
    fn.column_dimensions["A"].width = 52
    fn.column_dimensions["B"].width = 16
    fn.column_dimensions["C"].width = 50
    amt = f"Orders!$C$2:$C${last}"
    dflag = f"Orders!$D$2:$D${last}"
    res = [
        (1, "Q1 (25th percentile)", f'=IF(COUNT({amt})<4,"Need 4 amounts",QUARTILE({amt},1))', USD, "Formula: QUARTILE(Orders!C2:C100,1)"),
        (2, "Q3 (75th percentile)", f'=IF(COUNT({amt})<4,"Need 4 amounts",QUARTILE({amt},3))', USD, "Formula: QUARTILE(Orders!C2:C100,3)"),
        (3, "IQR (Q3 minus Q1)", '=IF(AND(ISNUMBER(B1),ISNUMBER(B2)),B2-B1,"")', USD, ""),
        (4, "Multiplier (type it)", 1.5, "0.0", "Typed. 1.5 is the usual rule; 3 flags only the far ones."),
        (5, "Amounts counted", f"=COUNT({amt})", "0", "Text and blank cells are not counted."),
        (6, "Upper fence (Q3 + multiplier x IQR)", '=IF(AND(ISNUMBER(B2),ISNUMBER(B3),ISNUMBER(B4)),B2+B4*B3,"Check")', USD, "Above this: Review."),
        (7, "Lower fence (Q1 - multiplier x IQR)", '=IF(AND(ISNUMBER(B1),ISNUMBER(B3),ISNUMBER(B4)),B1-B4*B3,"Check")', USD, "Below this: Review."),
        (8, "Rows flagged Review", f'=COUNTIF({dflag},"Review")', "0", "Look at the row. Do not delete it."),
        (9, "Rows OK", f'=COUNTIF({dflag},"OK")', "0", ""),
        (10, "Rows Check (text amount or no fences)", f'=COUNTIF({dflag},"Check")', "0", ""),
        (11, "Rows with an amount entry", f"=COUNTA({amt})", "0", ""),
        (12, "Check: Review + OK + Check = rows with an entry", '=IF(B8+B9+B10=B11,"OK","CHECK")', "General", ""),
        (14, "Mean plus or minus 2 standard deviations", None, "General", ""),
        (15, "Mean", f'=IF(COUNT({amt})<2,"",AVERAGE({amt}))', USD, ""),
        (16, "Standard deviation", f'=IF(COUNT({amt})<2,"",STDEV({amt}))', USD, ""),
        (17, "Upper limit (mean + 2 x SD)", '=IF(ISNUMBER(B15),B15+2*B16,"")', USD, ""),
        (18, "Lower limit (mean - 2 x SD)", '=IF(ISNUMBER(B15),B15-2*B16,"")', USD, "Can be below $0.00, so a tiny amount cannot fall under it."),
        (19, "Amounts outside those limits", f'=IF(ISNUMBER(B17),COUNTIF({amt},">"&B17)+COUNTIF({amt},"<"&B18),"")', "0", "Compare with Rows flagged Review above."),
    ]
    for r, label, v, fmt, note in res:
        fn[f"A{r}"] = label
        fn[f"A{r}"].font = Font(bold=True, size=12 if r == 14 else 11)
        if v is not None:
            fn[f"B{r}"] = v
            fn[f"B{r}"].number_format = fmt
            fn[f"B{r}"].fill = INPUT_FILL if r == 4 else FORMULA_FILL
            fn[f"B{r}"].alignment = Alignment(horizontal="right")
        if note:
            fn[f"C{r}"] = note
    fn.conditional_formatting.add("B12", FormulaRule(formula=['B12<>"OK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Find Unusually High or Low Order Amounts",
        "",
        "1. Orders tab: paste or type orders in A to C (yellow). Rows 2 to 100 have a Flag formula ready. The 30 sample orders are fictional; two are typing slips on purpose ($500.00 meant as $50.00 and $4.90 meant as $49.00).",
        "2. Fences tab: B1 is =QUARTILE(Orders!C2:C100,1) and B2 is =QUARTILE(Orders!C2:C100,3). B3 is the IQR, B2-B1. B4 is the multiplier you type (1.5).",
        "3. B6 = B2+B4*B3 is the upper fence and B7 = B1-B4*B3 is the lower fence. QUARTILE here is the inclusive method (same as QUARTILE.INC).",
        "4. Orders D2: =IF(C2=\"\",\"\",IF(NOT(ISNUMBER(C2)),\"Check\",IF(OR(C2>Fences!$B$6,C2<Fences!$B$7),\"Review\",\"OK\"))) plus a guard for missing fences. An amount equal to a fence is OK. Review means look at the row; it does not mean the amount is wrong.",
        "5. Rows flagged Review are shaded amber by a conditional formatting rule on the Flag text. A text amount (such as $50 typed as text) shows Check and is not counted in the fences.",
        "6. Fences B12 shows OK when Review + OK + Check equals the rows that have an amount entry.",
        "7. Fences B14 to B19 repeat the check with mean plus or minus 2 standard deviations. In the sample, the $500.00 typo raises the standard deviation so much that the lower limit falls below $0.00, and $4.90 cannot be caught by that rule (B19 counts 1, B8 counts 2). Compare B19 with B8 on your own data.",
        "8. With fewer than 4 numbers the fences show Need 4 amounts and every amount shows Check. If most amounts are equal, the IQR is 0 and any different amount is flagged.",
        "9. The fences are a screen for typing slips, not a rule about which orders are real or good. Check the row against the order itself.",
        "",
        "Orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Orders"].print_area = "A1:D41"
    wb.save(OUT / "tidy-tabs-outlier-order-amounts.xlsx")


def roas_calculator():
    wb = Workbook()
    ws = wb.active
    ws.title = "Campaigns"
    ws.column_dimensions["A"].width = 46
    for col, w in zip("BCDEFGH", [13, 10, 13, 10, 14, 16, 20]):
        ws.column_dimensions[col].width = w
    ws["A1"], ws["B1"] = "Gross margin before ads", 0.40
    ws["C1"] = "Typed: (price - product cost - fees) / price"
    ws["A2"], ws["B2"] = "Break even ROAS", '=IF(AND(ISNUMBER(B1),B1>0,B1<=1),1/B1,"Check margin")'
    ws["C2"] = "Formula: 1 / margin"
    ws["A1"].font = ws["A2"].font = Font(bold=True)
    ws["B1"].number_format = "0.0%"
    ws["B2"].number_format = "0.00"
    ws["B1"].fill = INPUT_FILL
    ws["B2"].fill = FORMULA_FILL
    for c in ("B1", "B2"):
        ws[c].alignment = Alignment(horizontal="right")
    for col, t in zip("ABCDEFGH", ["Campaign", "Ad spend", "Orders", "Sales", "ROAS", "Cost per order",
                                   "Profit after ads", "Status"]):
        c = ws[f"{col}4"]
        c.value, c.fill, c.font = t, HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 30
    camps = [
        ("Etsy Ads, candles", 180.00, 31, 1240.00),
        ("Facebook holiday gift sets", 420.00, 38, 1260.00),
        ("Google Shopping, soap", 250.00, 29, 725.00),
        ("Instagram reel boost", 150.00, 7, 245.00),
        ("Pinterest pins, wraps", 90.00, 12, 450.00),
        ("Etsy Ads, mugs", 120.00, 9, 300.00),
        ("Farmers market flyers", 60.00, 0, 0.00),
        ("Email sale announcement", 40.00, 22, 960.00),
    ]
    first, last = 5, 16
    for r in range(first, last + 1):
        if r - first < len(camps):
            for col, v in zip("ABCD", camps[r - first]):
                ws[f"{col}{r}"] = v
        ws[f"E{r}"] = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(D{r}),B{r}>0),D{r}/B{r},"")'
        ws[f"F{r}"] = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r}),C{r}>0),B{r}/C{r},"")'
        ws[f"G{r}"] = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(D{r}),ISNUMBER($B$2)),D{r}*$B$1-B{r},"")'
        ws[f"H{r}"] = (f'=IF(AND(B{r}="",C{r}="",D{r}=""),"",IF(NOT(ISNUMBER(G{r})),"Check",'
                       f'IF(B{r}=0,"No spend",IF(ROUND(G{r},2)>=0,"Above break-even","Below break-even"))))')
        for col, fmt in zip("BCD", [USD, "0", USD]):
            ws[f"{col}{r}"].number_format = fmt
            ws[f"{col}{r}"].fill = INPUT_FILL
        ws[f"A{r}"].fill = INPUT_FILL
        for col, fmt in zip("EFG", ["0.00", USD, USD]):
            ws[f"{col}{r}"].number_format = fmt
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"H{r}"].fill = FORMULA_FILL
    t = last + 1
    ws[f"A{t}"] = "Total"
    ws[f"B{t}"], ws[f"C{t}"], ws[f"D{t}"] = f"=SUM(B{first}:B{last})", f"=SUM(C{first}:C{last})", f"=SUM(D{first}:D{last})"
    ws[f"E{t}"] = f'=IF(SUM(B{first}:B{last})>0,SUM(D{first}:D{last})/SUM(B{first}:B{last}),"")'
    ws[f"F{t}"] = f'=IF(SUM(C{first}:C{last})>0,SUM(B{first}:B{last})/SUM(C{first}:C{last}),"")'
    ws[f"G{t}"] = f"=SUM(G{first}:G{last})"
    for col, fmt in zip("BCDEFG", [USD, "0", USD, "0.00", USD, USD]):
        ws[f"{col}{t}"].number_format = fmt
    for col in "ABCDEFGH":
        ws[f"{col}{t}"].font = Font(bold=True)
        ws[f"{col}{t}"].fill = FORMULA_FILL
    ws[f"A{t + 2}"] = "Check: total profit = Sales x margin - Spend"
    ws[f"A{t + 2}"].font = Font(bold=True)
    ws[f"B{t + 2}"] = (f'=IF(NOT(ISNUMBER($B$2)),"Check margin",IF(COUNTIF(H{first}:H{last},"Check")>0,"CHECK",'
                       f'IF(ABS(G{t}-(D{t}*$B$1-B{t}))<0.005,"OK","CHECK")))')
    ws[f"B{t + 2}"].fill = FORMULA_FILL
    ws[f"B{t + 2}"].alignment = Alignment(horizontal="right")
    red = PatternFill("solid", bgColor="F8D7DA", fgColor="F8D7DA")
    ws.conditional_formatting.add(f"A{first}:H{last}", FormulaRule(formula=[f'$H{first}="Below break-even"'], fill=red))
    ws.conditional_formatting.add(f"H{first}:H{last}", FormulaRule(formula=[f'$H{first}="Check"'], fill=ALERT_FILL))
    ws.conditional_formatting.add(f"B{t + 2}", FormulaRule(formula=[f'B{t + 2}<>"OK"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Return on Ad Spend (ROAS) and Break Even ROAS",
        "",
        "1. Campaigns tab, B1: your gross margin before ads as a percent, typed. Work it out from price, product cost and selling fees: (price - cost - fees) / price. The 40% here is a made-up example.",
        "2. B2 is the break even ROAS, =1/B1. At a 40% margin every $1.00 of ads must bring back $2.50 in sales just to cover the product cost, fees and the ad spend. A margin of 0 or below, above 100%, or text shows Check margin.",
        "3. Rows 5 to 16: one campaign per row. Type the ad spend (B), the orders (C) and the sales (D) from the ad report. They are yellow.",
        "4. ROAS E5 =D5/B5. Cost per order F5 =B5/C5. Profit after ads G5 =D5*$B$1-B5. A row with no spend leaves ROAS blank instead of showing #DIV/0!, and its status is No spend.",
        "5. Status H5: Above break-even when profit after ads rounds to $0.00 or more, Below break-even when it is less. Check means the spend or sales is blank or text, or the margin is not valid. Orders (C) are used only for cost per order.",
        "6. Total row 17: overall ROAS is total sales divided by total spend. B19 shows OK when total profit equals total sales x margin minus total spend, and CHECK when any row needs checking.",
        "7. This is arithmetic only. It does not say what a good ROAS is, and it ignores returns, shipping, taxes and the time between a click and a sale. Ad platforms count orders and sales in their own ways; follow the platform's report.",
        "",
        "Campaigns, spend and sales are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb["Campaigns"].print_title_rows = None
    wb.save(OUT / "tidy-tabs-roas-calculator.xlsx")


def pareto_top_products():
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    header(ws, ["Product", "Revenue", "Share of sales", "Cumulative share", "Group"], [26, 14, 14, 16, 10])
    prods = [
        ("Soy candle, 8 oz", 12480.00), ("Lavender soap bar", 9860.50), ("Holiday gift set", 7420.00),
        ("Beeswax food wrap", 5310.25), ("Linen tote bag", 3890.00), ("Ceramic mug", 2675.50),
        ("Wool dryer balls", 1980.00), ("Hand cream, 2 oz", 1450.75), ("Cedar sachet", 1120.00),
        ("Bath salts, 12 oz", 860.50), ("Lip balm 3-pack", 640.00), ("Wick trimmer", 455.25),
        ("Candle sampler", 310.00), ("Match box", 215.50), ("Soap dish", 142.00),
    ]
    last = 40
    tot = f"SUM($B$2:$B${last})"
    for r in range(2, last + 1):
        if r - 2 < len(prods):
            ws[f"A{r}"], ws[f"B{r}"] = prods[r - 2]
        ws[f"B{r}"].number_format = USD
        ws[f"C{r}"] = f'=IF(OR(B{r}="",NOT(ISNUMBER(B{r})),{tot}<=0),"",B{r}/{tot})'
        ws[f"D{r}"] = f'=IF(OR(B{r}="",NOT(ISNUMBER(B{r})),{tot}<=0),"",SUM($B$2:B{r})/{tot})'
        ws[f"E{r}"] = (f'=IF(B{r}="","",IF(OR(NOT(ISNUMBER(B{r})),{tot}<=0,NOT(ISNUMBER(Summary!$B$1))),"Check",'
                       f'IF(SUM($B$1:B{r - 1})/{tot}<Summary!$B$1,"Top","Rest")))')
        ws[f"C{r}"].number_format = ws[f"D{r}"].number_format = "0.0%"
        ws[f"A{r}"].fill = ws[f"B{r}"].fill = INPUT_FILL
        for col in "CDE":
            ws[f"{col}{r}"].fill = FORMULA_FILL
    green = PatternFill("solid", bgColor="E2F0D9", fgColor="E2F0D9")
    ws.conditional_formatting.add(f"A2:E{last}", FormulaRule(formula=['$E2="Top"'], fill=green))
    ws.conditional_formatting.add(f"E2:E{last}", FormulaRule(formula=['$E2="Check"'], fill=ALERT_FILL))

    sm = wb.create_sheet("Summary", 1)
    sm.column_dimensions["A"].width = 50
    sm.column_dimensions["B"].width = 18
    rev = f"Products!$B$2:$B${last}"
    grp = f"Products!$E$2:$E${last}"
    res = [
        ("Cut line (share of sales, type it)", 0.8, "0%"),
        ("Products listed", f"=COUNT({rev})", "0"),
        ("Products in the Top group", f'=COUNTIF({grp},"Top")', "0"),
        ("Share of products in Top", '=IF(B2>0,B3/B2,"")', "0.0%"),
        ("Share of revenue in Top", f'=IF(SUM({rev})>0,SUMIF({grp},"Top",{rev})/SUM({rev}),"")', "0.0%"),
        ("Sorted check", f'=IF(SUMPRODUCT(--(Products!$B$2:$B${last - 1}<Products!$B$3:$B${last}))=0,"Sorted","Sort largest first")', "General"),
        ("Rows to check (text or negative revenue)", f'=COUNTIF({grp},"Check")', "0"),
        ("Total revenue", f"=SUM({rev})", USD),
        ("Check: groups cover every product, shares add to 100%",
         f'=IF(NOT(ISNUMBER(B1)),"Check settings",IF(AND(B6="Sorted",B7=0,B2=B3+COUNTIF({grp},"Rest"),'
         f'ABS(SUM(Products!$C$2:$C${last})-1)<0.000001),"OK","CHECK"))', "General"),
    ]
    for i, (label, v, fmt) in enumerate(res, start=1):
        sm[f"A{i}"], sm[f"B{i}"] = label, v
        sm[f"A{i}"].font = Font(bold=True)
        sm[f"B{i}"].number_format = fmt
        sm[f"B{i}"].alignment = Alignment(horizontal="right")
        sm[f"B{i}"].fill = INPUT_FILL if i == 1 else FORMULA_FILL
    sm.conditional_formatting.add("B9", FormulaRule(formula=['B9<>"OK"'], fill=ALERT_FILL))
    sm.conditional_formatting.add("B6", FormulaRule(formula=['B6<>"Sorted"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Find the Products That Make Up 80 Percent of Sales",
        "",
        "1. Products tab: type or paste the product and its revenue (A and B, yellow). Sort the table largest revenue first: select A1:B40 and use Data > Sort on Revenue, largest to smallest. Rows 2 to 40 have formulas ready.",
        "2. Share C2 =B2/SUM($B$2:$B$40). Cumulative D2 =SUM($B$2:B2)/SUM($B$2:$B$40) is the running share of total revenue.",
        "3. Group E2 =IF(SUM($B$1:B1)/SUM($B$2:$B$40)<Summary!$B$1,\"Top\",\"Rest\"). A product is Top while the revenue above it is still under the cut line (80%), so the product that crosses 80% is the last one in Top. A product that lands on exactly 80.0% is Top, and the next one is Rest. SUM of the header row above row 2 is 0.",
        "4. Summary B1 is the cut line (yellow). B3 counts Top products, B4 is their share of the product count, B5 is their share of revenue (at or above the cut line).",
        "5. Summary B6 says Sorted or Sort largest first. If it says Sort largest first, the groups are wrong until you sort.",
        "6. A text revenue or a negative revenue (a refund) shows Check in Group. Text is left out of the total; a negative number is in the total, so fix it before reading the groups. Summary B9 shows OK only when sorted, nothing to check, and the shares add to 100%.",
        "7. The 80/20 split is a rule of thumb. Real shops rarely land exactly on 80 and 20. Use the Top group to decide what to restock or promote; the sheet does not say why a product sells.",
        "8. No chart is included. The In-cell bars template shows how to draw bars with REPT.",
        "",
        "Products and revenue are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Products"].print_area = "A1:E41"
    wb.save(OUT / "tidy-tabs-pareto-top-products.xlsx")


def _style_inputs(ws, cells):
    for c in cells:
        ws[c].fill = INPUT_FILL


def _ok_format(ws, cell):
    ws.conditional_formatting.add(cell, FormulaRule(formula=[f'AND({cell}<>"OK",{cell}<>"")'], fill=ALERT_FILL))


def annual_growth_rate():
    wb = Workbook()
    ws = wb.active
    ws.title = "Growth"
    for col, w in zip("ABCDEF", [10, 16, 3, 42, 16, 52]):
        ws.column_dimensions[col].width = w
    for col, t in zip("AB", ["Year", "Sales"]):
        c = ws[f"{col}1"]
        c.value, c.fill, c.font = t, HEADER_FILL, HEADER_FONT
    sales = [(2022, 12000.00), (2023, 14200.00), (2024, 15900.00), (2025, 19100.00), (2026, 21500.00)]
    for r in range(2, 12):
        if r - 2 < len(sales):
            ws[f"A{r}"], ws[f"B{r}"] = sales[r - 2]
        ws[f"A{r}"].fill = ws[f"B{r}"].fill = INPUT_FILL
        ws[f"A{r}"].number_format = "0"
        ws[f"B{r}"].number_format = USD
    rows = [
        (1, "Start year (typed)", 2022, "0", True, "Must match a year in column A."),
        (2, "End year (typed)", 2026, "0", True, ""),
        (3, "Periods (years between)", '=IF(AND(ISNUMBER(E1),ISNUMBER(E2)),E2-E1,"Check inputs")', "0", False,
         "End year minus start year, not the number of rows."),
        (4, "Start year sales", '=IF(ISNUMBER(E1),IFERROR(INDEX($B$2:$B$11,MATCH(E1,$A$2:$A$11,0)),"Check inputs"),"Check inputs")', USD, False, ""),
        (5, "End year sales", '=IF(ISNUMBER(E2),IFERROR(INDEX($B$2:$B$11,MATCH(E2,$A$2:$A$11,0)),"Check inputs"),"Check inputs")', USD, False, ""),
        (6, "Average yearly growth (RRI)",
         '=IF(OR(NOT(ISNUMBER(E3)),NOT(ISNUMBER(E4)),NOT(ISNUMBER(E5))),"Check inputs",'
         'IF(OR(E4<=0,E3<1,E5<0),"Check inputs",_xlfn.RRI(E3,E4,E5)))', "0.00%", False,
         "Formula: =RRI(E3,E4,E5), plus guards"),
        (7, "Same rate with the power form", '=IF(ISNUMBER(E6),(E5/E4)^(1/E3)-1,"Check inputs")', "0.00%", False,
         "Formula: =(E5/E4)^(1/E3)-1"),
        (8, "Total growth, first to last", '=IF(ISNUMBER(E6),E5/E4-1,"")', "0.00%", False, ""),
        (9, "End sales rebuilt from the rate", '=IF(ISNUMBER(E6),E4*(1+E6)^E3,"")', USD, False,
         "Formula: =E4*(1+E6)^E3"),
        (10, "Check: rebuilt end sales = real end sales",
         '=IF(NOT(ISNUMBER(E6)),"No result",IF(AND(ABS(E9-E5)<0.005,ABS(E6-E7)<0.0000001),"OK","CHECK"))',
         "General", False, "OK when both forms agree and the rebuild matches to the cent."),
    ]
    for r, label, v, fmt, inp, note in rows:
        ws[f"D{r}"] = label
        ws[f"D{r}"].font = Font(bold=True)
        ws[f"E{r}"] = v
        ws[f"E{r}"].number_format = fmt
        ws[f"E{r}"].fill = INPUT_FILL if inp else FORMULA_FILL
        ws[f"E{r}"].alignment = Alignment(horizontal="right")
        if note:
            ws[f"F{r}"] = note
    _ok_format(ws, "E10")
    ws.conditional_formatting.add("E6", FormulaRule(formula=['NOT(ISNUMBER(E6))'], fill=ALERT_FILL))

    ck = wb.create_sheet("Check")
    header(ck, ["Year", "Sales", "Change vs year before", "Sales at the average rate", "Real minus average rate"],
           [10, 16, 20, 24, 24])
    for r in range(2, 12):
        ck[f"A{r}"] = f'=IF(Growth!A{r}="","",Growth!A{r})'
        ck[f"B{r}"] = f'=IF(OR(Growth!A{r}="",Growth!B{r}=""),"",Growth!B{r})'
        if r == 2:
            ck[f"C{r}"] = ""
        else:
            ck[f"C{r}"] = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(B{r - 1})),IF(B{r - 1}>0,B{r}/B{r - 1}-1,""),"")'
        ck[f"D{r}"] = (f'=IF(AND(ISNUMBER(Growth!$E$6),ISNUMBER(A{r}),ISNUMBER(Growth!$E$1),ISNUMBER(Growth!$E$2)),'
                       f'IF(AND(A{r}>=Growth!$E$1,A{r}<=Growth!$E$2),IF(A{r}=Growth!$E$1,Growth!$E$4,Growth!$E$4*(1+Growth!$E$6)^(A{r}-Growth!$E$1)),""),"")')
        ck[f"E{r}"] = f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(B{r})),B{r}-D{r},"")'
        ck[f"A{r}"].number_format = "0"
        ck[f"B{r}"].number_format = ck[f"D{r}"].number_format = ck[f"E{r}"].number_format = USD
        ck[f"C{r}"].number_format = "0.0%"
        for col in "ABCDE":
            ck[f"{col}{r}"].fill = FORMULA_FILL
    ck["A13"], ck["B13"] = "Best single year", '=IF(COUNT(C2:C11)=0,"",MAX(C2:C11))'
    ck["A14"], ck["B14"] = "Weakest single year", '=IF(COUNT(C2:C11)=0,"",MIN(C2:C11))'
    ck["B13"].number_format = ck["B14"].number_format = "0.0%"
    ck["A13"].font = ck["A14"].font = Font(bold=True)

    notes_sheet(wb, [
        "Tidy Tabs: Average Yearly Sales Growth (CAGR) With RRI",
        "",
        "1. Growth tab: type each year in column A and its sales in column B (yellow, up to 10 years in rows 2 to 11). Type the start year in E1 and the end year in E2. Both must appear in column A.",
        "2. E3 is the number of periods, =E2-E1. Four years of growth between 2022 and 2026, not five: count the gaps, not the years listed.",
        "3. E4 and E5 pick the start and end sales with INDEX and MATCH on the year.",
        "4. E6 is =RRI(E3,E4,E5). RRI takes the number of periods, the start value and the end value and returns the growth rate per period. Excel and Google Sheets have RRI. In this file it is stored as _xlfn.RRI, which Excel and LibreOffice read as RRI.",
        "5. E7 gives the same answer with =(E5/E4)^(1/E3)-1. E8 is the total growth from first to last year. E9 grows the start sales at the E6 rate for E3 years; E10 shows OK when E9 equals the real end sales to the cent.",
        "6. A start of 0 or less, an end below 0, fewer than 1 period (same year, or years reversed), a year that is not in column A, or text in a year or sales cell shows Check inputs instead of #NUM!. An end of 0 gives -100.00%.",
        "7. Check tab: the change from each year to the one before, and what each year would be if sales grew at the average rate. One average rate hides uneven years: compare columns C and E.",
        "8. This is arithmetic on past sales. It is not a forecast and does not say what a good rate is.",
        "",
        "Years and sales are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Growth"].print_title_rows = None
    wb["Growth"].print_area = "A1:F12"
    wb["Check"].print_area = "A1:E14"
    wb.save(OUT / "tidy-tabs-annual-growth-rate.xlsx")


def order_size_ranges():
    from openpyxl.worksheet.formula import ArrayFormula
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    header(ws, ["Order ID", "Date", "Amount"], [12, 13, 12])
    amts = [12.00, 24.00, 36.50, 8.00, 52.00, 25.00, 18.50, 74.00, 4.00, 31.00,
            140.00, 22.00, 49.99, 60.00, 15.00, 28.50, 50.00, 9.00, 87.50, 19.99,
            33.00, 41.00, 100.00, 14.00, 26.00, 65.00, 11.00, 38.00, 125.00, 20.00,
            45.00, 6.50, 55.00, 17.00, 29.00, 72.00, 23.50, 98.00, 13.00, 35.00]
    assert len(amts) == 40
    start = date(2026, 9, 1).toordinal()
    for i in range(2, 201):
        if i - 2 < len(amts):
            ws[f"A{i}"] = f"ORD-{5001 + i - 2}"
            ws[f"B{i}"] = date.fromordinal(start + (i - 2))
            ws[f"C{i}"] = amts[i - 2]
        for col in "ABC":
            ws[f"{col}{i}"].fill = INPUT_FILL
        ws[f"B{i}"].number_format = US_DATE
        ws[f"C{i}"].number_format = USD
    ws.print_area = "A1:C41"

    rg = wb.create_sheet("Ranges", 1)
    heads = ["From (typed)", "Up to, not including", "Range", "Orders", "Sales in range", "Share of orders",
             "Share of sales"]
    for col, t, w in zip("ABCDEFG", heads, [14, 20, 26, 10, 16, 14, 14]):
        c = rg[f"{col}1"]
        c.value, c.fill, c.font = t, HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
        rg.column_dimensions[col].width = w
    rg.row_dimensions[1].height = 32
    amt = "Orders!$C$2:$C$200"
    limits = [0, 25, 50, 100]
    for k, r in enumerate(range(2, 6)):
        rg[f"A{r}"] = limits[k]
        rg[f"A{r}"].fill = INPUT_FILL
        rg[f"A{r}"].number_format = USD
        if r < 5:
            rg[f"B{r}"] = f"=A{r + 1}"
            rg[f"C{r}"] = f'=IF(AND(ISNUMBER(A{r}),ISNUMBER(B{r})),TEXT(A{r},"$#,##0.00")&" up to "&TEXT(B{r},"$#,##0.00"),"")'
            rg[f"D{r}"] = f'=COUNTIFS({amt},">="&A{r},{amt},"<"&B{r})'
            rg[f"E{r}"] = f'=SUMIFS({amt},{amt},">="&A{r},{amt},"<"&B{r})'
        else:
            rg[f"B{r}"] = "and up"
            rg[f"C{r}"] = f'=IF(ISNUMBER(A{r}),TEXT(A{r},"$#,##0.00")&" and up","")'
            rg[f"D{r}"] = f'=COUNTIFS({amt},">="&A{r})'
            rg[f"E{r}"] = f'=SUMIFS({amt},{amt},">="&A{r})'
        rg[f"B{r}"].number_format = USD
    rg["A6"], rg["B6"] = "Below", "=A2"
    rg["B6"].number_format = USD
    rg["C6"] = '=IF(ISNUMBER(A2),"Below "&TEXT(A2,"$#,##0.00"),"")'
    rg["D6"] = f'=COUNTIFS({amt},"<"&A2)'
    rg["E6"] = f'=SUMIFS({amt},{amt},"<"&A2)'
    rg["A7"] = "Total"
    rg["D7"], rg["E7"] = "=SUM(D2:D6)", "=SUM(E2:E6)"
    for r in range(2, 8):
        rg[f"F{r}"] = f'=IF($D$7>0,D{r}/$D$7,"")'
        rg[f"G{r}"] = f'=IF($E$7<>0,E{r}/$E$7,"")'
        rg[f"E{r}"].number_format = USD
        rg[f"F{r}"].number_format = rg[f"G{r}"].number_format = "0.0%"
        for col in "BCDEFG":
            rg[f"{col}{r}"].fill = FORMULA_FILL
    rg["A6"].fill = rg["A7"].fill = FORMULA_FILL
    for col in "ABCDEFG":
        rg[f"{col}7"].font = Font(bold=True)

    summ = [
        (9, "Numeric amounts on Orders", f"=COUNT({amt})", "0"),
        (10, "Text amounts (not counted)", f"=COUNTA({amt})-B9", "0"),
        (11, "Limits typed in order", '=IF(AND(COUNT(A2:A5)=4,A2<A3,A3<A4,A4<A5),"OK","Check limits")', "General"),
        (12, "Sales on Orders", f"=SUM({amt})", USD),
        (13, "Check: ranges add up to the orders list",
         '=IF(AND(B11="OK",B10=0,D7=B9,ABS(E7-B12)<0.005,COUNTIF(E18:E22,"OK")=5),"OK","CHECK")', "General"),
    ]
    for r, label, f, fmt in summ:
        rg[f"A{r}"] = label
        rg[f"A{r}"].font = Font(bold=True)
        rg[f"B{r}"] = f
        rg[f"B{r}"].number_format = fmt
        rg[f"B{r}"].fill = FORMULA_FILL
        rg[f"B{r}"].alignment = Alignment(horizontal="right")
    rg.column_dimensions["A"].width = 38
    _ok_format(rg, "B13")
    rg.conditional_formatting.add("B11", FormulaRule(formula=['B11<>"OK"'], fill=ALERT_FILL))
    rg["A15"] = "Optional cross-check with FREQUENCY (array formula in C18:C22; amounts in whole cents)"
    rg["A15"].font = Font(bold=True)
    for col, t in zip("ABCDE", ["Bin top (limit minus $0.01)", "Range", "FREQUENCY", "COUNTIFS", "Same?"]):
        c = rg[f"{col}17"]
        c.value, c.fill, c.font = t, HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    rg.row_dimensions[17].height = 32
    maps = [(18, "A2", "C6", "D6"), (19, "A3", "C2", "D2"), (20, "A4", "C3", "D3"), (21, "A5", "C4", "D4")]
    for r, lim, lab, cnt in maps:
        rg[f"A{r}"] = f'=IF(ISNUMBER({lim}),ROUND({lim}-0.01,2),0)'
        rg[f"B{r}"] = f"={lab}"
        rg[f"D{r}"] = f"={cnt}"
        rg[f"A{r}"].number_format = USD
    rg["B22"], rg["D22"] = "=C5", "=D5"
    rg["C18"] = ArrayFormula("C18:C22", f'=IF(COUNT({amt})=0,0,IFERROR(FREQUENCY({amt},A18:A21),""))')
    for r in range(18, 23):
        rg[f"E{r}"] = f'=IF(C{r}=D{r},"OK","Differs")'
        for col in "ABCDE":
            rg[f"{col}{r}"].fill = FORMULA_FILL
    rg.conditional_formatting.add("E18:E22", FormulaRule(formula=['E18="Differs"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Count Orders by Dollar Amount Range With COUNTIFS",
        "",
        "1. Orders tab: one row per order, amount in column C (yellow, rows 2 to 200). Amounts only, no $ typed as text.",
        "2. Ranges tab, A2:A5: type the lower limit of each range (yellow). Keep them in ascending order. B2:B4 take the next row's limit; the last range has no upper limit.",
        "3. Orders in a range: =COUNTIFS(Orders!$C$2:$C$200,\">=\"&A2,Orders!$C$2:$C$200,\"<\"&B2). The range includes its lower limit and stops just before the upper one, so an order of exactly $25.00 counts in the range that starts at $25.00, never in two ranges.",
        "4. Sales in the range use SUMIFS with the same two tests (column E). Shares in F and G are the range divided by the total.",
        "5. Row 6 counts everything below the first limit, such as a $0.00 order or a negative refund, so no order is lost. Row 7 is the total.",
        "6. B13 shows OK when the limits are in order, no amount is text, the counts add up to the numeric amounts on Orders, the sales add up to the sum of the Orders column, and the FREQUENCY cross-check agrees with COUNTIFS in every range.",
        "7. The FREQUENCY cross-check (C18:C22) is an array formula. FREQUENCY counts values up to and including each bin, so each bin top is the limit minus $0.01, which assumes amounts in whole cents. Excel older than 365 needs Ctrl+Shift+Enter for array formulas; the cross-check is optional.",
        "8. COUNTIFS skips text amounts and blank cells. A text amount makes B10 more than 0 and B13 shows CHECK.",
        "9. Counts only: this says how many orders fall in each range in this sample, not what a good mix is.",
        "",
        "Orders and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Ranges"].print_title_rows = None
    wb.save(OUT / "tidy-tabs-order-size-ranges.xlsx")


def price_by_effective_date():
    wb = Workbook()
    pr = wb.active
    pr.title = "Prices"
    header(pr, ["Product", "Effective date", "New price", "Row check"], [26, 15, 12, 14])
    hist = [
        ("Cold Brew Pack", date(2026, 1, 1), 36.00), ("Cold Brew Pack", date(2026, 6, 1), 38.00),
        ("Cold Brew Pack", date(2026, 10, 1), 39.50),
        ("Dark Roast", date(2025, 1, 1), 43.00), ("Dark Roast", date(2026, 3, 1), 45.50),
        ("Decaf Blend", date(2025, 1, 1), 48.00), ("Decaf Blend", date(2025, 9, 1), 49.50),
        ("Decaf Blend", date(2026, 3, 1), 51.00), ("Decaf Blend", date(2026, 9, 1), 52.50),
        ("Espresso Blend", date(2025, 1, 1), 45.00), ("Espresso Blend", date(2025, 7, 1), 46.75),
        ("House Blend", date(2025, 1, 1), 42.00), ("House Blend", date(2025, 7, 1), 44.50),
        ("House Blend", date(2026, 3, 1), 46.00),
        ("Single Origin Ethiopia", date(2025, 6, 1), 58.00), ("Single Origin Ethiopia", date(2026, 4, 1), 61.00),
    ]
    for r in range(2, 41):
        if r - 2 < len(hist):
            pr[f"A{r}"], pr[f"B{r}"], pr[f"C{r}"] = hist[r - 2]
        pr[f"D{r}"] = (f'=IF(AND(A{r}="",B{r}="",C{r}=""),"",IF(NOT(AND(A{r}<>"",ISNUMBER(B{r}),ISNUMBER(C{r}))),"Check",'
                       f'IF(AND(A{r - 1}=A{r},ISNUMBER(B{r - 1})),IF(B{r}<B{r - 1},"Check order",IF(B{r}=B{r - 1},"Same day","OK")),"OK")))')
        pr[f"B{r}"].number_format = US_DATE
        pr[f"C{r}"].number_format = USD
        for col in "ABC":
            pr[f"{col}{r}"].fill = INPUT_FILL
        pr[f"D{r}"].fill = FORMULA_FILL
    pr.conditional_formatting.add("D2:D40", FormulaRule(formula=['OR(D2="Check",D2="Check order")'], fill=ALERT_FILL))
    pr.print_area = "A1:D41"

    od = wb.create_sheet("Orders", 1)
    header(od, ["Order date", "Product", "Qty", "Price in effect", "Line total"], [13, 26, 8, 16, 14])
    orders = [
        (date(2025, 12, 15), "Dark Roast", 4), (date(2025, 1, 1), "Dark Roast", 2),
        (date(2025, 6, 30), "House Blend", 6), (date(2025, 7, 1), "House Blend", 6),
        (date(2026, 2, 28), "Dark Roast", 3), (date(2026, 3, 1), "Dark Roast", 3),
        (date(2026, 3, 1), "House Blend", 5), (date(2026, 2, 27), "House Blend", 5),
        (date(2025, 3, 15), "Single Origin Ethiopia", 2), (date(2025, 6, 1), "Single Origin Ethiopia", 2),
        (date(2026, 4, 2), "Single Origin Ethiopia", 3), (date(2026, 1, 1), "Cold Brew Pack", 10),
        (date(2025, 11, 20), "Cold Brew Pack", 8), (date(2026, 7, 14), "Cold Brew Pack", 12),
        (date(2026, 10, 1), "Cold Brew Pack", 12), (date(2025, 8, 31), "Decaf Blend", 4),
        (date(2025, 9, 1), "Decaf Blend", 4), (date(2026, 8, 31), "Decaf Blend", 5),
        (date(2026, 9, 1), "Decaf Blend", 5), (date(2026, 10, 12), "Decaf Blend", 6),
        (date(2025, 5, 5), "Espresso Blend", 7), (date(2025, 7, 15), "Espresso Blend", 7),
        (date(2026, 10, 9), "Espresso Blend", 9), (date(2025, 2, 10), "House Blend", 3),
        (date(2025, 9, 9), "House Blend", 8), (date(2026, 5, 18), "House Blend", 8),
        (date(2025, 4, 22), "Dark Roast", 5), (date(2026, 9, 30), "Dark Roast", 6),
        (date(2026, 8, 15), "Cold Brew Pack", 6), (date(2026, 6, 1), "Cold Brew Pack", 4),
    ]
    assert len(orders) == 30
    P = "Prices!"
    for r in range(2, 201):
        if r - 2 < len(orders):
            od[f"A{r}"], od[f"B{r}"], od[f"C{r}"] = orders[r - 2]
        od[f"D{r}"] = (f'=IF(AND(A{r}="",B{r}=""),"",IF(OR(NOT(ISNUMBER(A{r})),B{r}=""),"Check",'
                       f'IFERROR(LOOKUP(2,1/(({P}$A$2:$A$40=B{r})*({P}$B$2:$B$40<=A{r})*ISNUMBER({P}$B$2:$B$40)*ISNUMBER({P}$C$2:$C$40)),{P}$C$2:$C$40),"No price yet")))')
        od[f"E{r}"] = f'=IF(ISNUMBER(D{r}),IF(ISNUMBER(C{r}),C{r}*D{r},"Check"),"")'
        od[f"A{r}"].number_format = US_DATE
        od[f"D{r}"].number_format = od[f"E{r}"].number_format = USD
        for col in "ABC":
            od[f"{col}{r}"].fill = INPUT_FILL
        od[f"D{r}"].fill = od[f"E{r}"].fill = FORMULA_FILL
    od.conditional_formatting.add("D2:E200", FormulaRule(formula=['OR(D2="Check",D2="No price yet",E2="Check")'], fill=ALERT_FILL))
    od.print_area = "A1:E41"

    od.column_dimensions["G"].width = 3
    od.column_dimensions["H"].width = 40
    od.column_dimensions["I"].width = 16
    for col in "HI":
        od[f"{col}1"].fill = HEADER_FILL
        od[f"{col}1"].font = HEADER_FONT
    od["H1"], od["I1"] = "Summary", "Value"
    rows_used = 'SUMPRODUCT(--((A2:A200<>"")+(B2:B200<>"")>0))'
    summ = [
        (2, "Orders priced", "=COUNT(E2:E200)", "0"),
        (3, "Orders with No price yet", '=COUNTIF(D2:D200,"No price yet")', "0"),
        (4, "Order rows to check", '=COUNTIF(D2:D200,"Check")+COUNTIF(E2:E200,"Check")', "0"),
        (5, "Price rows to check", '=COUNTIF(Prices!D2:D40,"Check")+COUNTIF(Prices!D2:D40,"Check order")', "0"),
        (6, "Total of line totals", "=SUM(E2:E200)", USD),
        (7, "Check: every order row is priced, flagged or No price yet",
         f'=IF(AND({rows_used}=COUNT(E2:E200)+COUNTIF(E2:E200,"Check")+COUNTIF(D2:D200,"No price yet")+COUNTIF(D2:D200,"Check"),I5=0),"OK","CHECK")',
         "General"),
    ]
    for r, label, f, fmt in summ:
        od[f"H{r}"] = label
        od[f"H{r}"].font = Font(bold=True)
        od[f"I{r}"] = f
        od[f"I{r}"].number_format = fmt
        od[f"I{r}"].fill = FORMULA_FILL
        od[f"I{r}"].alignment = Alignment(horizontal="right")
    _ok_format(od, "I7")
    od["H9"] = "One product with INDEX and MATCH (Cold Brew Pack, Prices rows 2 to 4)"
    od["H9"].font = Font(bold=True)
    od["H10"], od["I10"] = "Order date (typed)", date(2026, 8, 15)
    od["I10"].number_format = US_DATE
    od["I10"].fill = INPUT_FILL
    od["H11"] = "Price in effect (INDEX and MATCH)"
    od["I11"] = '=IF(NOT(ISNUMBER(I10)),"Check",IFERROR(INDEX(Prices!$C$2:$C$4,MATCH(I10,Prices!$B$2:$B$4,1)),"No price yet"))'
    od["H12"] = "Dates sorted oldest first?"
    od["I12"] = '=IF(AND(Prices!B2<=Prices!B3,Prices!B3<=Prices!B4),"Sorted","Not sorted")'
    od["I11"].number_format = USD
    for c in ("I11", "I12"):
        od[c].fill = FORMULA_FILL
        od[c].alignment = Alignment(horizontal="right")
    od.conditional_formatting.add("I12", FormulaRule(formula=['I12<>"Sorted"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Look Up the Price in Effect on a Date",
        "",
        "1. Prices tab: one row for each price change. Product (A), effective date (B, typed as MM/DD/YYYY) and the new price (C) are yellow. Keep each product's rows together, oldest date first. Rows 2 to 40 are read.",
        "2. Orders tab: order date (A), product (B) and quantity (C) are yellow, rows 2 to 200.",
        "3. D2 is =IFERROR(LOOKUP(2,1/((Prices!$A$2:$A$40=B2)*(Prices!$B$2:$B$40<=A2)),Prices!$C$2:$C$40),\"No price yet\"). It tests every price row: right product and an effective date on or before the order date. The template adds ISNUMBER(Prices!$B$2:$B$40)*ISNUMBER(Prices!$C$2:$C$40) inside the 1/(...) so a row with a text date or a blank or text price is skipped (the Prices tab marks that row Check). LOOKUP(2,1/...) returns the last row where all the tests are true, so it works without pressing Ctrl+Shift+Enter and the rows do not have to be sorted. If two rows for the same product have the same date, the lower row wins.",
        "4. An order dated exactly on an effective date gets the new price. The day before gets the old price. A date before the product's first price, or a product that is not in the list, shows No price yet instead of #N/A. A text date or an empty product shows Check.",
        "5. E2 is the line total, =C2*D2. A blank or text quantity shows Check.",
        "6. Prices column D checks each row: Check means a missing product, a date or price that is not a number; Check order means a date earlier than the row above for the same product; Same day means the same date as the row above (the lower row wins).",
        "7. Orders I7 shows OK when every order row is priced, flagged Check, or No price yet, and no Prices row needs checking.",
        "8. H9:I12 shows the INDEX and MATCH version for one product: =INDEX(Prices!$C$2:$C$4,MATCH(I10,Prices!$B$2:$B$4,1)). MATCH with 1 finds the last date that is not after the order date, but it needs the dates sorted oldest first (I12 says Sorted). With unsorted dates it can return a wrong price without any error.",
        "9. Product names must match exactly (a trailing space is a different product). This is a price lookup only: no discounts, tax or shipping.",
        "",
        "Products, dates and prices are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb.save(OUT / "tidy-tabs-price-by-effective-date.xlsx")


def batches_from_materials():
    wb = Workbook()
    bt = wb.active
    bt.title = "Batch"
    bt.column_dimensions["A"].width = 34
    bt.column_dimensions["B"].width = 12
    bt.column_dimensions["C"].width = 56
    bt["A1"], bt["B1"], bt["C1"] = "Candles one batch makes", 12, "Typed. Any unit you count the finished product in."
    bt["A2"], bt["B2"], bt["C2"] = "Batches to plan for (N)", 10, "Typed. Used for the to-buy column on the Materials tab."
    for c in ("A1", "A2"):
        bt[c].font = Font(bold=True)
    _style_inputs(bt, ["B1", "B2"])

    ws = wb.create_sheet("Materials", 1)
    header(ws, ["Material", "On hand", "Needed per batch", "Whole batches possible", "Left after the batches made",
                "To buy to make N batches"], [26, 11, 15, 16, 18, 18])
    mats = [("Soy wax (lb)", 42.5, 4.5), ("Fragrance oil (oz)", 30, 3.5), ("Cotton wicks", 130, 12),
            ("8 oz jars", 87, 12), ("Labels", 150, 12), ("Shipping boxes", 100, 12)]
    for r in range(2, 8):
        ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = mats[r - 2]
        ws[f"D{r}"] = (f'=IF(AND(B{r}="",C{r}=""),"",IF(OR(NOT(ISNUMBER(B{r})),AND(C{r}<>"",NOT(ISNUMBER(C{r})))),"Check",'
                       f'IF(OR(B{r}<0,C{r}<0),"Check",IF(OR(C{r}="",C{r}=0),"",INT(ROUND(B{r}/C{r},10))))))')
        ws[f"E{r}"] = f'=IF(AND(ISNUMBER(D{r}),ISNUMBER($B$10)),ROUND(B{r}-C{r}*$B$10,6),"")'
        ws[f"F{r}"] = f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(Batch!$B$2)),ROUND(MAX(0,C{r}*Batch!$B$2-B{r}),6),"")'
        for col in "ABC":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "DEF":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"B{r}"].number_format = ws[f"C{r}"].number_format = "#,##0.##"
        ws[f"D{r}"].number_format = "0"
        ws[f"E{r}"].number_format = ws[f"F{r}"].number_format = "#,##0.##"
    ws["A9"] = "Results"
    ws["A9"].font = Font(bold=True, size=12)
    res = [
        (10, "Batches you can make", '=IF(COUNTIF(D2:D7,"Check")>0,"Check inputs",IF(COUNT(D2:D7)=0,"No materials",MIN(D2:D7)))', "0"),
        (11, "Material that runs out first", '=IF(ISNUMBER(B10),INDEX(A2:A7,MATCH(B10,D2:D7,0))&"","")', "General"),
        (12, "Units made", '=IF(AND(ISNUMBER(B10),ISNUMBER(Batch!B1)),B10*Batch!B1,"")', "#,##0.##"),
        (13, "Batches to plan for (N)", '=IF(ISNUMBER(Batch!B2),Batch!B2,"")', "#,##0.##"),
        (14, "Materials to buy for N batches", '=IF(COUNT(F2:F7)=0,"",COUNTIF(F2:F7,">0"))', "0"),
        (15, "Check: no negative leftover, limit under one batch",
         '=IF(NOT(ISNUMBER(B10)),"No result",IF(AND(MIN(E2:E7)>=0,'
         'INDEX(E2:E7,MATCH(B10,D2:D7,0))<INDEX(C2:C7,MATCH(B10,D2:D7,0))),"OK","CHECK"))', "General"),
    ]
    for r, label, f, fmt in res:
        ws[f"A{r}"] = label
        ws[f"A{r}"].font = Font(bold=True)
        ws[f"B{r}"] = f
        ws[f"B{r}"].number_format = fmt
        ws[f"B{r}"].fill = FORMULA_FILL
        ws[f"B{r}"].alignment = Alignment(horizontal="right")
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 18
    _ok_format(ws, "B15")
    ws.conditional_formatting.add("A2:F7", FormulaRule(formula=['AND(ISNUMBER($D2),ISNUMBER($B$10),$D2=$B$10)'], fill=ALERT_FILL))
    ws.conditional_formatting.add("D2:D7", FormulaRule(formula=['D2="Check"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: How Many Batches From the Materials on Hand",
        "",
        "1. Batch tab: type how many finished items one batch makes (B1) and how many batches you want to plan for (B2). Both are yellow.",
        "2. Materials tab, rows 2 to 7: the material name (A), how much you have (B) and how much one batch needs (C). Use the same unit in B and C for a row (pounds with pounds, ounces with ounces). They are yellow.",
        "3. D2 is the whole batches that material allows: INT(B2/C2), written as INT(ROUND(B2/C2,10)), which is INT(B2/C2) with a guard against tiny binary rounding errors in decimal divisions such as 0.3 / 0.1. Rows with a blank or 0 in C are left out (no #DIV/0!). Text or a negative number shows Check.",
        "4. B10 is =MIN(D2:D7): the whole batches you can make. B11 names the material behind it with INDEX and MATCH. If two materials tie, the first one listed is shown. B12 is B10 times the batch yield.",
        "5. Column E is what is left of each material after B10 batches (B2-C2*$B$10). Column F is what to buy to make N batches in total, =MAX(0,C2*N-B2), where N is Batch!B2. The row that limits the batches is shaded.",
        "6. B15 shows OK when no leftover is negative and the limiting material has less than one batch left.",
        "7. Whole batches only. The sheet does not know recipes, yields, shelf life or safety rules. It only divides what you have by what you use.",
        "",
        "Materials and amounts are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb)
    wb["Materials"].print_title_rows = None
    wb.save(OUT / "tidy-tabs-batches-from-materials.xlsx")


def appointment_schedule():
    from datetime import time
    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule"
    heads = ["Client", "Service", "Start time", "Minutes", "End time", "Gap to next (min)", "Status",
             "Start (min of day)", "End (min of day)"]
    header(ws, heads, [16, 20, 12, 9, 12, 12, 16, 12, 12])
    ws.row_dimensions[1].height = 32
    for c in ws[1]:
        c.alignment = Alignment(vertical="center", wrap_text=True)
    appts = [
        ("Bailey", "Full groom", time(9, 0), 90), ("Mochi", "Bath and brush", time(10, 45), 45),
        ("Cooper", "Full groom", time(11, 30), 90), ("Luna", "Nail trim", time(13, 15), 30),
        ("Rosie", "Bath and brush", time(14, 0), 60), ("Max", "Full groom", time(15, 10), 90),
        ("Daisy", "Nail trim", time(16, 45), 30), ("Buddy", "Puppy trim", time(17, 30), 45),
    ]
    for r in range(2, 13):
        if r - 2 < len(appts):
            for col, v in zip("ABCD", appts[r - 2]):
                ws[f"{col}{r}"] = v
        ws[f"E{r}"] = f'=IF(OR(NOT(ISNUMBER(C{r})),NOT(ISNUMBER(D{r}))),"",C{r}+D{r}/1440)'
        ws[f"F{r}"] = f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(C{r + 1})),ROUND((C{r + 1}-E{r})*1440,0),"")'
        ws[f"G{r}"] = (
            f'=IF(AND(A{r}="",C{r}="",D{r}=""),"",IF(OR(NOT(ISNUMBER(C{r})),NOT(ISNUMBER(D{r}))),"Check",'
            f'IF(OR(C{r}<0,C{r}>=1,D{r}<=0,D{r}<>INT(D{r}),NOT(ISNUMBER($L$1))),"Check",IF(E{r}>=1,"Past midnight",'
            f'IF(F{r}="",IF(AND(A{r + 1}="",C{r + 1}="",D{r + 1}=""),"Last","Check next row"),'
            f'IF(C{r + 1}<C{r},"Out of order",IF(F{r}<0,"Overlap",IF(F{r}<$L$1,"Short buffer","OK"))))))))')
        ws[f"H{r}"] = f'=IF(ISNUMBER(C{r}),ROUND(C{r}*1440,0),"")'
        ws[f"I{r}"] = f'=IF(ISNUMBER(E{r}),ROUND(E{r}*1440,0),"")'
        for col in "ABCD":
            ws[f"{col}{r}"].fill = INPUT_FILL
        for col in "EFGHI":
            ws[f"{col}{r}"].fill = FORMULA_FILL
        ws[f"C{r}"].number_format = ws[f"E{r}"].number_format = "h:mm AM/PM"
        ws[f"D{r}"].number_format = ws[f"F{r}"].number_format = ws[f"H{r}"].number_format = ws[f"I{r}"].number_format = "0"
    ws.column_dimensions["J"].width = 3
    ws.column_dimensions["K"].width = 46
    ws.column_dimensions["L"].width = 14
    ws["K1"], ws["L1"] = "Buffer between appointments (min)", 15
    ws["L1"].fill = INPUT_FILL
    tot = [
        (2, "Booked minutes", "=SUM(D2:D12)", "0"),
        (3, "First start", '=IF(COUNT(C2:C12)=0,"",MIN(C2:C12))', "h:mm AM/PM"),
        (4, "Last end", '=IF(COUNT(E2:E12)=0,"",MAX(E2:E12))', "h:mm AM/PM"),
        (5, "Open minutes between", '=IF(AND(ISNUMBER(L3),ISNUMBER(L4)),ROUND((L4-L3)*1440,0)-L2,"")', "0"),
        (6, "Appointments", "=COUNT(C2:C12)", "0"),
        (7, "Rows to look at", '=COUNTIF(G2:G12,"Short buffer")+COUNTIF(G2:G12,"Overlap")', "0"),
        (8, "Check: booked + gaps = first start to last end",
         '=IF(COUNT(C2:C12)=0,"No appointments",IF(COUNTIF(G2:G12,"Overlap")+COUNTIF(G2:G12,"Past midnight")'
         '+COUNTIF(G2:G12,"Check")+COUNTIF(G2:G12,"Check next row")+COUNTIF(G2:G12,"Out of order")>0,"Fix rows first",'
         'IF(L2+SUM(F2:F12)=ROUND((L4-L3)*1440,0),"OK","CHECK")))', "General"),
    ]
    for r, label, f, fmt in tot:
        ws[f"K{r}"] = label
        ws[f"L{r}"] = f
        ws[f"L{r}"].number_format = fmt
        ws[f"L{r}"].fill = FORMULA_FILL
    for r in range(1, 9):
        ws[f"K{r}"].font = Font(bold=True)
        ws[f"L{r}"].alignment = Alignment(horizontal="right")
    _ok_format(ws, "L8")
    amber = PatternFill("solid", bgColor="FFE699", fgColor="FFE699")
    ws.conditional_formatting.add("A2:G12", FormulaRule(formula=['$G2="Short buffer"'], fill=amber))
    ws.conditional_formatting.add("A2:G12", FormulaRule(
        formula=['OR($G2="Overlap",$G2="Past midnight",$G2="Check",$G2="Check next row",$G2="Out of order")'],
        fill=ALERT_FILL))

    sl = wb.create_sheet("Slots", 1)
    sl.column_dimensions["A"].width = 24
    for col, w in zip("BCDEFG", [14, 22, 12, 3, 22, 10]):
        sl.column_dimensions[col].width = w
    sl["A1"], sl["B1"] = "Opening time", time(9, 0)
    sl["A2"], sl["B2"] = "Slot length (min)", 30
    sl["A3"], sl["B3"] = "Closing time", time(18, 30)
    sl["B1"].number_format = sl["B3"].number_format = "h:mm AM/PM"
    for c in ("A1", "A2", "A3"):
        sl[c].font = Font(bold=True)
    _style_inputs(sl, ["B1", "B2", "B3"])
    for col, t in zip("ABCD", ["Slot start", "Slot end", "Appointments touching the slot", "Status"]):
        c = sl[f"{col}4"]
        c.value, c.fill, c.font = t, HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    sl.row_dimensions[4].height = 32
    okslot = 'IFERROR(AND(ISNUMBER($B$1),ISNUMBER($B$2),ISNUMBER($B$3),$B$2>0,$B$2=INT($B$2),$B$3>$B$1),FALSE)'
    for r in range(5, 53):
        if r == 5:
            sl[f"A{r}"] = f'=IF(AND({okslot},$B$1>=0,$B$1<1),ROUND($B$1*1440,0)/1440,"")'
        else:
            sl[f"A{r}"] = f'=IF(ISNUMBER(A{r - 1}),IF(AND(ROUND(A{r - 1}*1440,0)+$B$2<1440,ROUND(A{r - 1}*1440,0)+$B$2<ROUND($B$3*1440,0)),(ROUND(A{r - 1}*1440,0)+$B$2)/1440,""),"")'
        sl[f"B{r}"] = f'=IF(ISNUMBER(A{r}),(ROUND(A{r}*1440,0)+$B$2)/1440,"")'
        sl[f"C{r}"] = (f'=IF(ISNUMBER(A{r}),COUNTIFS(Schedule!$H$2:$H$12,"<"&ROUND(B{r}*1440,0),'
                       f'Schedule!$I$2:$I$12,">"&ROUND(A{r}*1440,0)),"")')
        sl[f"D{r}"] = f'=IF(C{r}="","",IF(C{r}=0,"Open",IF(C{r}=1,"Taken","Shared slot")))'
        sl[f"A{r}"].number_format = sl[f"B{r}"].number_format = "h:mm AM/PM"
        sl[f"C{r}"].number_format = "0"
        for col in "ABCD":
            sl[f"{col}{r}"].fill = FORMULA_FILL
    sl["F4"], sl["G4"] = "Summary", "Value"
    sl["F4"].fill = sl["G4"].fill = HEADER_FILL
    sl["F4"].font = sl["G4"].font = HEADER_FONT
    for r, label, st in [(5, "Open slots", "Open"), (6, "Taken slots", "Taken"), (7, "Shared slots", "Shared slot")]:
        sl[f"F{r}"] = label
        sl[f"G{r}"] = f'=COUNTIF(D5:D52,"{st}")'
        sl[f"F{r}"].font = Font(bold=True)
        sl[f"G{r}"].fill = FORMULA_FILL
    sl.conditional_formatting.add("A5:D52", FormulaRule(formula=['$D5="Taken"'], fill=DONE_FILL))
    sl.conditional_formatting.add("A5:D52", FormulaRule(formula=['$D5="Shared slot"'], fill=ALERT_FILL))

    notes_sheet(wb, [
        "Tidy Tabs: Appointment Schedule With End Times, Gaps and Buffers",
        "",
        "1. Schedule tab: type the client (A), service (B), start time (C) and length in whole minutes (D). They are yellow. Type times as 9:00 AM and list the day in start-time order, rows 2 to 12. L1 is the buffer you want between appointments, in minutes.",
        "2. End time, E2 =C2+D2/1440. A day has 24 x 60 = 1,440 minutes, so minutes divided by 1440 is the part of a day to add. Format the result as a time.",
        "3. Gap to next, F2 =ROUND((C3-E2)*1440,0): the next start minus this end, times 1440, rounded to whole minutes. The ROUND stops times stored as fractions of a day from showing 14.999999 instead of 15.",
        "4. Status, G2: Short buffer when the gap is under L1, Overlap when it is negative, OK otherwise. Last is the final appointment. Past midnight means the end time is on the next day (the time shown wraps around); the sheet does not plan across days. Check means a start that is not a time of day, minutes that are not a whole number above 0, text, or a buffer in L1 that is not a number. Out of order means the next row starts earlier.",
        "5. Totals, L2:L8: booked minutes, first start, last end, open minutes between (first start to last end minus booked). L8 shows OK when booked minutes plus the gaps equal the time from first start to last end. It says Fix rows first while a row is Overlap, Past midnight, Check or Out of order, and CHECK if a blank row sits in the middle of the list.",
        "6. Columns H and I turn start and end into whole minutes of the day (9:00 AM is 540). The Slots tab compares those whole numbers so floating point does not move an appointment into the wrong slot.",
        "7. Slots tab: type the opening time (B1), the slot length in whole minutes (B2) and the closing time (B3). A slot is Taken when one appointment starts before the slot ends and ends after the slot starts. Shared slot means two appointments touch the same slot (they can be back to back inside it). Open means none. Slots stop at the closing time; the table holds up to 48 slots.",
        "8. Times and minutes only. No booking policy, deposit or pay rule.",
        "",
        "Clients, pets and times are fictional. Checked in LibreOffice Calc only; not opened in Excel or Google Sheets.",
        "Google Sheets: File > Import > Upload this .xlsx.",
    ])
    finish_letter(wb, landscape=True)
    wb["Schedule"].print_title_rows = None
    wb["Schedule"].print_area = "A1:L13"
    wb["Slots"].print_title_rows = None
    wb["Slots"].page_setup.orientation = "portrait"
    wb.save(OUT / "tidy-tabs-appointment-schedule.xlsx")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    inventory()
    projects()
    zip_codes()
    print_one_page()
    split_city_state_zip()
    dependent_dropdown()
    client_contact_list()
    packing_slip()
    running_balance()
    highlight_duplicates()
    business_days_ship_by()
    time_to_decimal_hours()
    receivables_aging()
    markup_vs_margin()
    vlookup_price_list()
    break_even_calculator()
    subscription_renewal_tracker()
    budget_vs_actual()
    farmers_market_sales_log()
    clean_customer_list()
    sale_price_discount()
    shipping_weight_tier_lookup()
    reorder_point_calculator()
    freelance_hourly_rate()
    convert_text_to_numbers()
    sales_tax_calculator()
    sum_expenses_by_category()
    round_prices_nearest_nickel_99()
    remove_non_breaking_spaces()
    year_to_date_sales_total()
    quantity_price_tiers()
    weighted_average_cost()
    last_order_date_maxifs()
    count_orders_by_month()
    average_order_value_by_channel()
    subtotal_filtered_sales()
    rolling_average_sales()
    prorate_partial_month()
    two_way_rate_card()
    in_cell_bars()
    quarterly_sales_totals()
    missing_invoice_numbers()
    sumif_contains_text()
    invoice_installments()
    marketplace_net_payout()
    invoice_due_dates()
    sales_by_weekday()
    repeat_customers()
    supplier_price_compare()
    split_sku_code()
    listing_length_check()
    ounces_to_pounds_ounces()
    last_number_in_column()
    time_zone_converter()
    count_weekdays_in_month()
    days_to_get_paid()
    early_payment_discount()
    sell_through_rate()
    boxes_needed()
    median_order_value()
    week_ending_date()
    free_shipping_threshold()
    forecast_next_month()
    client_tenure()
    incomplete_rows()
    overlapping_bookings()
    cash_drawer_count()
    nth_weekday_of_month()
    format_phone_numbers()
    scale_recipe_batch()
    lookup_two_criteria()
    stock_on_hand_log()
    email_domain_split()
    on_time_shipping_rate()
    price_per_square_foot()
    expiring_dates()
    validation_rules()
    sum_with_or()
    fill_blanks_above()
    mailing_labels()
    months_to_pay_off()
    sum_across_monthly_sheets()
    outlier_order_amounts()
    roas_calculator()
    pareto_top_products()
    annual_growth_rate()
    order_size_ranges()
    price_by_effective_date()
    batches_from_materials()
    appointment_schedule()
    print("\n".join(sorted(p.name for p in OUT.iterdir())))
