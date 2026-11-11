"""Budget Tracker: generates Budget_Tracker.xlsx with Budget, Transactions, and Dashboard tabs.
Run: pip install openpyxl  then  python budget_tracker.py
Edit the `cats` list and income values below to customize."""
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference

F = "Arial"
BLUE = Font(name=F, color="0000FF"); BLACK = Font(name=F); GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True); TITLE = Font(name=F, bold=True, size=14)
HDR = Font(name=F, bold=True, color="FFFFFF"); HFILL = PatternFill("solid", fgColor="305496")
YEL = PatternFill("solid", fgColor="FFFF00"); TOT = PatternFill("solid", fgColor="D9E1F2")
NOTE = Font(name=F, italic=True, color="595959", size=9)
USD = '$#,##0;($#,##0);-'; PCT = '0.0%'
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)

def hdr(ws, row, labels):
    for i, l in enumerate(labels, 1):
        c = ws.cell(row, i, l); c.font = HDR; c.fill = HFILL; c.alignment = Alignment(horizontal="center")

wb = Workbook()

# ---------------- Budget ----------------
b = wb.active; b.title = "Budget"
b["A1"] = "Monthly Budget Plan"; b["A1"].font = TITLE
b["A2"] = "Edit the yellow cells (blue text). Black cells are formulas. Green cells pull from another sheet."; b["A2"].font = NOTE

b["A4"] = "Income"; b["A4"].font = BOLD
rows = [("Gross Annual Salary", 125000, USD, "Starting salary, computer programmer in San Diego (from class template)"),
        ("Estimated Tax Rate", 0.3027, PCT, "From template: $37,838 tax on $125,000"),]
for i, (lbl, v, fmt, note) in enumerate(rows, 5):
    b.cell(i, 1, lbl).font = BLACK
    c = b.cell(i, 2, v); c.font = BLUE; c.fill = YEL; c.number_format = fmt
    b.cell(i, 4, note).font = NOTE
calc = [("Gross Monthly Income", "=B5/12"), ("Monthly Tax", "=B7*B6"), ("Net Monthly Income", "=B7-B8")]
for i, (lbl, f) in enumerate(calc, 7):
    b.cell(i, 1, lbl).font = BOLD if i == 9 else BLACK
    c = b.cell(i, 2, f); c.font = BOLD if i == 9 else BLACK; c.number_format = USD

hdr(b, 11, ["Category", "Group", "Type", "Monthly Budget", "Annual", "Notes"])
cats = [
 ("Rent", "Fixed", "Need", 2357, "Independent studio apartment"),
 ("Insurance (Auto)", "Fixed", "Need", 200, "Small SUV"),
 ("Insurance (Medical)", "Fixed", "Need", 250, "Estimated"),
 ("Utilities", "Variable", "Need", 120, "Electricity to trash"),
 ("Cell Phone", "Variable", "Need", 75, "T-Mobile mid-tier"),
 ("Car Gas", "Variable", "Need", 300, "Website + calculator"),
 ("House Gas", "Variable", "Need", 20, "Estimated"),
 ("Car Repair", "Variable", "Need", 200, "Estimated"),
 ("Internet", "Variable", "Need", 35, "AT&T"),
 ("Groceries", "Variable", "Need", 200, "One vegetarian person"),
 ("Medical", "Variable", "Need", 60, "Estimated"),
 ("Household", "Variable", "Need", 60, "Estimated"),
 ("Eating Out", "Variable", "Want", 200, "Professional outings"),
 ("Clothing", "Variable", "Want", 180, "Estimated"),
 ("Cosmetic", "Variable", "Want", 45, "From experience"),
 ("Recreation/Entertainment", "Variable", "Want", 50, "Estimated"),
 ("Gifts (family)", "Variable", "Want", 623, "Fixed: annual is now 623 x 12"),
 ("Short-Term Savings", "Savings", "Savings", 726, "Emergency fund, big purchases"),
 ("Long-Term Savings", "Savings", "Savings", 1563, "401k and other plans"),
]
start = 12
for i, (cat, grp, typ, amt, note) in enumerate(cats, start):
    for col, v in ((1, cat), (2, grp), (3, typ)):
        c = b.cell(i, col, v); c.font = BLUE; c.fill = YEL
    c = b.cell(i, 4, amt); c.font = BLUE; c.fill = YEL; c.number_format = USD
    c = b.cell(i, 5, f"=D{i}*12"); c.font = BLACK; c.number_format = USD
    b.cell(i, 6, note).font = NOTE
end = start + len(cats) - 1
R = f"${start}:${end}"
dv = DataValidation(type="list", formula1='"Fixed,Variable,Savings"'); b.add_data_validation(dv); dv.add(f"B{start}:B{end}")
dv2 = DataValidation(type="list", formula1='"Need,Want,Savings"'); b.add_data_validation(dv2); dv2.add(f"C{start}:C{end}")

r = end + 2
tots = [("Total Fixed", f'=SUMIFS($D${start}:$D${end},$B${start}:$B${end},"Fixed")'),
        ("Total Variable", f'=SUMIFS($D${start}:$D${end},$B${start}:$B${end},"Variable")'),
        ("Total Savings", f'=SUMIFS($D${start}:$D${end},$B${start}:$B${end},"Savings")'),
        ("Total Budgeted", f"=SUM(D{start}:D{end})"),
        ("Left to Assign (Net - Budgeted)", f"=B9-D{r+3}")]
for k, (lbl, f) in enumerate(tots):
    rr = r + k
    b.cell(rr, 1, lbl).font = BOLD
    c = b.cell(rr, 4, f); c.font = BOLD; c.number_format = USD; c.fill = TOT
    c = b.cell(rr, 5, f"=D{rr}*12"); c.font = BOLD; c.number_format = USD; c.fill = TOT
b.cell(r + 4, 6, "Aim for $0: every dollar of net income has a job").font = NOTE
left_row = r + 4
b.conditional_formatting.add(f"D{left_row}", CellIsRule(operator="lessThan", formula=["-1"], font=Font(name=F, bold=True, color="C00000")))

r = left_row + 2
b.cell(r, 1, "50/30/20 Check").font = BOLD
hdr(b, r + 1, ["Type", "Monthly", "% of Net", "Target", "Difference"])
for k, (typ, tgt) in enumerate((("Need", 0.5), ("Want", 0.3), ("Savings", 0.2))):
    rr = r + 2 + k
    b.cell(rr, 1, typ).font = BLACK
    c = b.cell(rr, 2, f'=SUMIFS($D${start}:$D${end},$C${start}:$C${end},A{rr})'); c.number_format = USD; c.font = BLACK
    c = b.cell(rr, 3, f"=IFERROR(B{rr}/$B$9,0)"); c.number_format = PCT; c.font = BLACK
    c = b.cell(rr, 4, tgt); c.number_format = PCT; c.font = BLUE; c.fill = YEL
    c = b.cell(rr, 5, f"=C{rr}-D{rr}"); c.number_format = '+0.0%;-0.0%;-'; c.font = BLACK
for col, w in zip("ABCDEF", (32, 14, 12, 16, 14, 46)): b.column_dimensions[col].width = w
b.freeze_panes = "A12"

# ---------------- Transactions ----------------
t = wb.create_sheet("Transactions")
t["A1"] = "Transactions"; t["A1"].font = TITLE
t["A2"] = "Log every purchase here. Pick the category from the dropdown so it matches the Budget sheet. Rows 5-9 are examples; delete them when you start."; t["A2"].font = NOTE
hdr(t, 4, ["Date", "Description", "Category", "Amount", "Notes"])
samples = [(date(2026,11,1), "November rent", "Rent", 2357), (date(2026,11,3), "Trader Joe's", "Groceries", 64.20),
           (date(2026,11,5), "Shell gas station", "Car Gas", 52.10), (date(2026,11,8), "Lunch with team", "Eating Out", 38.75),
           (date(2026,11,10), "Movie tickets", "Recreation/Entertainment", 31.00)]
TROWS = 500
for i in range(5, 5 + TROWS):
    t.cell(i, 1).number_format = "mm/dd/yyyy"; t.cell(i, 4).number_format = '$#,##0.00'
    for col in range(1, 6): t.cell(i, col).font = BLUE
for i, (d, desc, cat, amt) in enumerate(samples, 5):
    t.cell(i, 1, d); t.cell(i, 2, desc); t.cell(i, 3, cat); t.cell(i, 4, amt); t.cell(i, 5, "Example")
dv3 = DataValidation(type="list", formula1=f"=Budget!$A${start}:$A${end}", allow_blank=True)
t.add_data_validation(dv3); dv3.add(f"C5:C{4+TROWS}")
for col, w in zip("ABCDE", (13, 32, 26, 13, 30)): t.column_dimensions[col].width = w
t.freeze_panes = "A5"

# ---------------- Dashboard ----------------
d = wb.create_sheet("Dashboard")
d["A1"] = "Budget vs. Actual"; d["A1"].font = TITLE
d["A3"] = "Month to view (any date in the month):"; d["A3"].font = BOLD
c = d["C3"]; c.value = date(2026, 11, 1); c.number_format = "mmmm yyyy"; c.font = BLUE; c.fill = YEL
d["E3"] = '=DATE(YEAR(C3),MONTH(C3),1)'; d["E3"].number_format = "mm/dd/yyyy"; d["E3"].font = Font(name=F, color="FFFFFF")
hdr(d, 5, ["Category", "Budgeted", "Actual", "Remaining", "% Used"])
T = f"Transactions!$D$5:$D${4+TROWS}"; TC = f"Transactions!$C$5:$C${4+TROWS}"; TD = f"Transactions!$A$5:$A${4+TROWS}"
for k in range(len(cats)):
    rr = 6 + k; br = start + k
    d.cell(rr, 1, f"=Budget!A{br}").font = GREEN
    c = d.cell(rr, 2, f"=Budget!D{br}"); c.font = GREEN; c.number_format = USD
    c = d.cell(rr, 3, f'=SUMIFS({T},{TC},A{rr},{TD},">="&$E$3,{TD},"<"&EDATE($E$3,1))'); c.font = BLACK; c.number_format = '$#,##0.00'
    c = d.cell(rr, 4, f"=B{rr}-C{rr}"); c.font = BLACK; c.number_format = '$#,##0.00;($#,##0.00);-'
    c = d.cell(rr, 5, f"=IFERROR(C{rr}/B{rr},0)"); c.font = BLACK; c.number_format = PCT
dend = 6 + len(cats) - 1; tr = dend + 1
d.cell(tr, 1, "Total").font = BOLD
for col, f in ((2, f"=SUM(B6:B{dend})"), (3, f"=SUM(C6:C{dend})"), (4, f"=B{tr}-C{tr}"), (5, f"=IFERROR(C{tr}/B{tr},0)")):
    c = d.cell(tr, col, f); c.font = BOLD; c.fill = TOT
    c.number_format = PCT if col == 5 else '$#,##0.00;($#,##0.00);-'
red = PatternFill("solid", fgColor="F8CBAD")
d.conditional_formatting.add(f"E6:E{tr}", CellIsRule(operator="greaterThan", formula=["1"], fill=red, font=Font(name=F, bold=True, color="C00000")))
d.conditional_formatting.add(f"D6:D{tr}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=F, bold=True, color="C00000")))
d.cell(tr + 2, 1, "Red = over budget for that category this month.").font = NOTE
d.cell(tr + 3, 1, "Savings rows show what you actually transferred to savings; log transfers as transactions too.").font = NOTE

ch = BarChart(); ch.type = "bar"; ch.title = "Budgeted vs. Actual"; ch.height = 11; ch.width = 18
ch.add_data(Reference(d, min_col=2, max_col=3, min_row=5, max_row=dend), titles_from_data=True)
ch.set_categories(Reference(d, min_col=1, min_row=6, max_row=dend))
d.add_chart(ch, "G5")
for col, w in zip("ABCDE", (28, 14, 14, 14, 10)): d.column_dimensions[col].width = w
d.freeze_panes = "A6"

wb.move_sheet("Dashboard", offset=-2)
wb.active = 0
wb.save("Budget_Tracker.xlsx")
