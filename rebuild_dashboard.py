import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Protection
import shutil

SRC = "/root/.claude/uploads/ca8257f0-1581-55bf-88a4-34c1d47ea869/fb5c87e5-Aura_Andermatt_10Year_Model.xlsx"
DST = "/home/user/andermatt-financial-model/Aura_Andermatt_10Year_Model.xlsx"

shutil.copy2(SRC, DST)
wb = openpyxl.load_workbook(DST)
ws = wb["Dashboard"]

# ── 1. Rename "Net Result" → "Profit After Tax" ──
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
    for cell in row:
        if isinstance(cell.value, str) and "Net Result" in cell.value:
            cell.value = cell.value.replace("Net Result", "Profit After Tax")

wb["P&L"]["A29"].value = "PROFIT AFTER TAX"

# ── 2. Remove "10-Year Total Revenue" row from KEY METRICS ──
for c in range(1, 5):
    ws.cell(row=6, column=c).value = None

# ── 3. Replace charts: remove all 3, add 2 new ones ──
ws._charts.clear()

# Categories (year headers) - row 11 has years 2027-2036
cats = Reference(ws, min_col=2, max_col=11, min_row=11)

NAVY = "1F4E79"
RED = "C62828"
GREEN = "2E7D32"

def make_simple_line_chart(ws, title, series_defs, cats):
    chart = LineChart()
    chart.title = title
    chart.style = 2
    chart.height = 12
    chart.width = 20
    chart.legend.position = 'b'

    chart.y_axis.numFmt = '#,##0'
    chart.y_axis.majorGridlines = None
    chart.y_axis.minorGridlines = None
    chart.y_axis.title = None
    chart.x_axis.title = None
    chart.x_axis.majorGridlines = None
    chart.x_axis.minorGridlines = None

    for row_num, label, color, dash in series_defs:
        ref = Reference(ws, min_col=2, max_col=11, min_row=row_num)
        chart.add_data(ref, from_rows=True)
        idx = len(chart.series) - 1
        s = chart.series[idx]
        s.title = openpyxl.chart.series.SeriesLabel(v=label)
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 25000
        if dash:
            s.graphicalProperties.line.dashStyle = dash

    chart.set_categories(cats)
    return chart

# Chart 1: Profit After Tax
chart1 = make_simple_line_chart(ws, "Profit After Tax", [
    (19, "Base", NAVY, None),
    (31, "Downside", RED, "dash"),
    (43, "Upside", GREEN, "dash"),
], cats)

# Chart 2: Cumulative Cash
chart2 = make_simple_line_chart(ws, "Cumulative Cash", [
    (20, "Base", NAVY, None),
    (32, "Downside", RED, "dash"),
    (44, "Upside", GREEN, "dash"),
], cats)

ws.add_chart(chart1, "M1")
ws.add_chart(chart2, "M17")

# ── 4. Sheet protection (password: 2027) ──
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
