import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Protection, PatternFill, Font
from openpyxl.comments import Comment
import shutil

SRC = "/root/.claude/uploads/ca8257f0-1581-55bf-88a4-34c1d47ea869/fb5c87e5-Aura_Andermatt_10Year_Model.xlsx"
DST = "/home/user/andermatt-financial-model/Aura_Andermatt_10Year_Model.xlsx"

shutil.copy2(SRC, DST)
wb = openpyxl.load_workbook(DST)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# REVENUE SHEET CHANGES
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
rev = wb["Revenue"]

# ── 1. Rename "GLOBAL REVENUE VARIABLES" → "MODEL ASSUMPTIONS" ──
rev["A3"].value = "MODEL ASSUMPTIONS"
ASSUMPTIONS_FILL = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")
ASSUMPTIONS_FONT = Font(bold=True, size=11, color="1F4E79")
rev["A3"].fill = ASSUMPTIONS_FILL
rev["A3"].font = ASSUMPTIONS_FONT

YELLOW = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
SECTION_FILL = PatternFill(start_color="FFD9E2F3", end_color="FFD9E2F3", fill_type="solid")
SECTION_FONT = Font(bold=True, size=11)
COMMENT_FONT = Font(size=9, color="666666")
LABEL_FONT = Font(size=11)

COL_LETTERS = ["", "A","B","C","D","E","F","G","H","I","J","K"]
YEARS = list(range(2027, 2037))

def clear_rows(ws, start_row, end_row):
    for r in range(start_row, end_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font()
            cell.comment = None

def write_year_headers(ws, row):
    for i, year in enumerate(YEARS):
        ws.cell(row=row, column=i+2).value = year
        ws.cell(row=row, column=i+2).font = Font(bold=True)

def write_section_header(ws, row, title):
    ws.cell(row=row, column=1).value = title
    ws.cell(row=row, column=1).fill = SECTION_FILL
    ws.cell(row=row, column=1).font = SECTION_FONT

def write_comment_row(ws, row, text):
    ws.cell(row=row, column=1).value = text
    ws.cell(row=row, column=1).font = COMMENT_FONT

def write_yellow_cell(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    cell.value = value
    cell.fill = YELLOW

def write_fb_block(ws, base_row, label_prefix, guest_count, per_person, share_rate):
    """Write a 5-row F&B sub-block: guest count, spend, share, est total, revenue."""
    r = base_row
    ws.cell(row=r, column=1).value = f"  {label_prefix} — Guest Count"
    ws.cell(row=r, column=1).font = LABEL_FONT
    write_yellow_cell(ws, r, 2, guest_count)

    r += 1
    ws.cell(row=r, column=1).value = f"  {label_prefix} — Per-person Spend"
    ws.cell(row=r, column=1).font = LABEL_FONT
    write_yellow_cell(ws, r, 2, per_person)

    r += 1
    ws.cell(row=r, column=1).value = f"  {label_prefix} — Share Rate"
    ws.cell(row=r, column=1).font = LABEL_FONT
    write_yellow_cell(ws, r, 2, share_rate)
    share_comment = Comment(
        "Share rate may be based on profit margin, not gross revenue. "
        "Adjust when contract terms are confirmed.",
        "Model Note"
    )
    share_comment.width = 300
    share_comment.height = 80
    ws.cell(row=r, column=2).comment = share_comment

    r += 1
    guest_row = base_row
    spend_row = base_row + 1
    ws.cell(row=r, column=1).value = f"  {label_prefix} — Est. Total Sales (reference)"
    ws.cell(row=r, column=1).font = LABEL_FONT
    ws.cell(row=r, column=2).value = f"=B{guest_row}*B{spend_row}"
    for i in range(1, 10):
        cl = COL_LETTERS[i+2]
        prev = COL_LETTERS[i+1]
        ws.cell(row=r, column=i+2).value = f"={prev}{r}*(1+$B$9)"

    r += 1
    est_row = r - 1
    share_row = base_row + 2
    ws.cell(row=r, column=1).value = f"  {label_prefix} — Aura Revenue"
    ws.cell(row=r, column=1).font = LABEL_FONT
    for i in range(10):
        cl = COL_LETTERS[i+2]
        ws.cell(row=r, column=i+2).value = f"={cl}{est_row}*$B{share_row}"

    return r  # last row used

# ── 2. Rewrite Bar/F&B section (rows 64+) ──
# Clear from row 64 to end of sheet
clear_rows(rev, 64, rev.max_row)

# Extend sheet if needed (original max was 101, new max is 114)
# Just writing to new rows will extend it automatically.

# Section header
ROW = 64
write_section_header(rev, ROW, "SECTION 4: BAR / F&B REVENUE")
ROW += 1
write_comment_row(rev, ROW, "  Share rates are currently set to 0%. Aura does not take a share of bar/food revenue. If a share agreement is negotiated, change the share rate.")
ROW += 2  # skip blank row
write_year_headers(rev, ROW)  # row 67
ROW += 1

# Friday Night: Coffee + Drinks (rows 68-72)
FB_FRIDAY_BASE = ROW  # 68
last = write_fb_block(rev, ROW, "Friday Night: Coffee + Drinks", 150, 50, 0)
FRIDAY_REV_ROW = last  # 72
ROW = last + 2  # skip blank

# Saturday Day: Drinks (rows 74-78)
FB_SATDAY_DRINKS_BASE = ROW  # 74
last = write_fb_block(rev, ROW, "Saturday Day: Drinks", 280, 20, 0)
SATDAY_DRINKS_REV_ROW = last  # 78
ROW = last + 2

# Saturday Day: Food (rows 80-84)
FB_SATDAY_FOOD_BASE = ROW  # 80
last = write_fb_block(rev, ROW, "Saturday Day: Food", 280, 50, 0)
SATDAY_FOOD_REV_ROW = last  # 84
ROW = last + 2

# Saturday Night: Drinks (rows 86-90)
FB_SATNIGHT_BASE = ROW  # 86
last = write_fb_block(rev, ROW, "Saturday Night: Drinks", 150, 50, 0)
SATNIGHT_REV_ROW = last  # 90
ROW = last + 2

# Sunday Coffee (rows 92-96)
FB_SUNDAY_BASE = ROW  # 92
last = write_fb_block(rev, ROW, "Sunday Coffee", 80, 8, 0)
SUNDAY_REV_ROW = last  # 96
ROW = last + 2

# Bar/F&B Total (row 98)
FB_TOTAL_ROW = ROW  # 98
rev.cell(row=ROW, column=1).value = "Bar/F&B Total"
rev.cell(row=ROW, column=1).font = Font(bold=True)
for i in range(10):
    cl = COL_LETTERS[i+2]
    rev.cell(row=ROW, column=i+2).value = (
        f"={cl}{FRIDAY_REV_ROW}+{cl}{SATDAY_DRINKS_REV_ROW}+"
        f"{cl}{SATDAY_FOOD_REV_ROW}+{cl}{SATNIGHT_REV_ROW}+{cl}{SUNDAY_REV_ROW}"
    )
ROW += 2

# ── Dinner Revenue ──
DINNER_HEADER_ROW = ROW  # 100
write_section_header(rev, ROW, "SECTION 5: DINNER REVENUE")
ROW += 1
write_year_headers(rev, ROW)  # 101
ROW += 1
DINNER_ROW = ROW  # 102
rev.cell(row=ROW, column=1).value = "Dinner Revenue (manual input)"
rev.cell(row=ROW, column=1).font = LABEL_FONT
for i in range(10):
    write_yellow_cell(rev, ROW, i+2, 0)
ROW += 1
write_comment_row(rev, ROW, "  BP: 'Dinner is invite-only, sponsor-funded, never sold.' If a dinner-specific sponsor appears, it goes under Partnerships.")
ROW += 2

# ── Other Revenue ──
OTHER_HEADER_ROW = ROW  # 105
write_section_header(rev, ROW, "SECTION 6: OTHER REVENUE")
ROW += 1
write_year_headers(rev, ROW)  # 106
ROW += 1
OTHER_ROW = ROW  # 107
rev.cell(row=ROW, column=1).value = "Other Revenue"
rev.cell(row=ROW, column=1).font = LABEL_FONT
write_yellow_cell(rev, ROW, 2, 2000)
for i in range(1, 10):
    cl = COL_LETTERS[i+2]
    prev = COL_LETTERS[i+1]
    rev.cell(row=ROW, column=i+2).value = f"={prev}{OTHER_ROW}*(1+'P&L'!$B$4)"
ROW += 1
write_comment_row(rev, ROW, "  Other Revenue grows with General Inflation Rate (P&L sheet). Contents: add-ons, upselling, merchandise.")
ROW += 2

# ── Revenue Totals ──
TOTALS_HEADER_ROW = ROW  # 110
rev.cell(row=ROW, column=1).value = "REVENUE TOTALS"
rev.cell(row=ROW, column=1).fill = SECTION_FILL
rev.cell(row=ROW, column=1).font = SECTION_FONT
ROW += 1
write_year_headers(rev, ROW)  # 111
ROW += 1

WINTER_ROW = ROW  # 112
rev.cell(row=ROW, column=1).value = "TOTAL WINTER REVENUE"
rev.cell(row=ROW, column=1).font = Font(bold=True)
for i in range(10):
    cl = COL_LETTERS[i+2]
    rev.cell(row=ROW, column=i+2).value = (
        f"={cl}40+{cl}56+{cl}61+{cl}{FB_TOTAL_ROW}+{cl}{DINNER_ROW}+{cl}{OTHER_ROW}"
    )
ROW += 1

SUMMER_ROW = ROW  # 113
rev.cell(row=ROW, column=1).value = "Summer Revenue"
rev.cell(row=ROW, column=1).font = Font(bold=True)
for i, year in enumerate(YEARS):
    cl = COL_LETTERS[i+2]
    rev.cell(row=ROW, column=i+2).value = f"=IF({year}>=$B$13,{cl}{WINTER_ROW}*$B$11,0)"
ROW += 1

ANNUAL_ROW = ROW  # 114
rev.cell(row=ROW, column=1).value = "TOTAL ANNUAL REVENUE"
rev.cell(row=ROW, column=1).font = Font(bold=True)
for i in range(10):
    cl = COL_LETTERS[i+2]
    rev.cell(row=ROW, column=i+2).value = f"={cl}{WINTER_ROW}+{cl}{SUMMER_ROW}"

# ── 3. Update P&L cross-references ──
pl = wb["P&L"]

# P&L R13: Bar/F&B Total (was Revenue!85 → now FB_TOTAL_ROW)
for i in range(10):
    cl = COL_LETTERS[i+2]
    pl.cell(row=13, column=i+2).value = f"=Revenue!{cl}{FB_TOTAL_ROW}"

# P&L R14: Dinner Revenue (was Revenue!89 → now DINNER_ROW)
for i in range(10):
    cl = COL_LETTERS[i+2]
    pl.cell(row=14, column=i+2).value = f"=Revenue!{cl}{DINNER_ROW}"

# P&L R15: Other Revenue (was Revenue!94 → now OTHER_ROW)
for i in range(10):
    cl = COL_LETTERS[i+2]
    pl.cell(row=15, column=i+2).value = f"=Revenue!{cl}{OTHER_ROW}"

# P&L R17: TOTAL ANNUAL REVENUE (was Revenue!101 → now ANNUAL_ROW)
for i in range(10):
    cl = COL_LETTERS[i+2]
    pl.cell(row=17, column=i+2).value = f"=Revenue!{cl}{ANNUAL_ROW}"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# COST SHEET CHANGES
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
cost = wb["Cost"]

# ── 1. Rename "GLOBAL COST VARIABLES" → "MODEL ASSUMPTIONS" ──
cost["A3"].value = "MODEL ASSUMPTIONS"
cost["A3"].fill = ASSUMPTIONS_FILL
cost["A3"].font = ASSUMPTIONS_FONT

# ── 2. Add TOTAL COST row after Founder Pay (row 55) ──
COST_TOTAL_ROW = 57
cost.cell(row=COST_TOTAL_ROW, column=1).value = "TOTAL COST"
cost.cell(row=COST_TOTAL_ROW, column=1).fill = SECTION_FILL
cost.cell(row=COST_TOTAL_ROW, column=1).font = SECTION_FONT
for i in range(10):
    cl = COL_LETTERS[i+2]
    cost.cell(row=COST_TOTAL_ROW, column=i+2).value = f"={cl}49+{cl}53+{cl}55"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DASHBOARD CHANGES (same as before)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ws = wb["Dashboard"]

# ── Rename "Net Result" → "Profit After Tax" ──
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
    for cell in row:
        if isinstance(cell.value, str) and "Net Result" in cell.value:
            cell.value = cell.value.replace("Net Result", "Profit After Tax")

pl["A29"].value = "PROFIT AFTER TAX"

# ── Remove "10-Year Total Revenue" row from KEY METRICS ──
for c in range(1, 5):
    ws.cell(row=6, column=c).value = None

# ── Replace charts: remove all 3, add 2 new ones ──
ws._charts.clear()

cats = Reference(ws, min_col=2, max_col=11, min_row=11)

NAVY = "1F4E79"
RED = "C62828"
GREEN = "2E7D32"

def make_chart(ws, title, series_defs, cats):
    chart = LineChart()
    chart.title = title
    chart.style = 2
    chart.height = 13
    chart.width = 22
    chart.legend.position = 'b'

    chart.y_axis.numFmt = '#,##0'
    chart.y_axis.title = "CHF"
    chart.x_axis.title = None
    chart.y_axis.minorGridlines = None

    for row_num, label, color, dash, width in series_defs:
        ref = Reference(ws, min_col=2, max_col=11, min_row=row_num)
        chart.add_data(ref, from_rows=True)
        idx = len(chart.series) - 1
        s = chart.series[idx]
        s.title = openpyxl.chart.series.SeriesLabel(v=label)
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = width
        if dash:
            s.graphicalProperties.line.dashStyle = dash

        s.dLbls = DataLabelList()
        s.dLbls.showVal = True
        s.dLbls.numFmt = '#,##0'
        s.dLbls.showCatName = False
        s.dLbls.showSerName = False

    chart.set_categories(cats)
    return chart

chart1 = make_chart(ws, "Profit After Tax (CHF)", [
    (19, "Base", NAVY, None, 28000),
    (31, "Downside", RED, "dash", 18000),
    (43, "Upside", GREEN, "dash", 18000),
], cats)

chart2 = make_chart(ws, "Cumulative Cash (CHF)", [
    (20, "Base", NAVY, None, 28000),
    (32, "Downside", RED, "dash", 18000),
    (44, "Upside", GREEN, "dash", 18000),
], cats)

ws.add_chart(chart1, "M1")
ws.add_chart(chart2, "M17")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SHEET PROTECTION (password: 2027)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PASSWORD = "2027"

for sheet_name in ["Dashboard", "P&L", "Revenue", "Cost"]:
    sheet = wb[sheet_name]
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell in row:
            cell.protection = Protection(locked=True)
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell in row:
            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                rgb = str(cell.fill.start_color.rgb)
                if rgb in ("FFFFFF00", "00FFFF00"):
                    cell.protection = Protection(locked=False)
    sheet.protection.sheet = True
    sheet.protection.password = PASSWORD
    sheet.protection.enable()
    sheet.protection.selectLockedCells = False
    sheet.protection.selectUnlockedCells = False

wb.save(DST)
print("Done! Saved to", DST)
print(f"New row mappings: FB Total={FB_TOTAL_ROW}, Dinner={DINNER_ROW}, Other={OTHER_ROW}, Annual={ANNUAL_ROW}")
