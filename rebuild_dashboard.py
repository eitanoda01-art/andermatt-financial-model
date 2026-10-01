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

# Chart 1: Profit After Tax (Line)
chart1 = LineChart()
chart1.title = "Profit After Tax (10-Year)"
chart1.style = 2
chart1.y_axis.title = "CHF"
chart1.y_axis.numFmt = '#,##0'
chart1.height = 12
chart1.width = 20
chart1.legend.position = 'b'
chart1.y_axis.crossesAt = 0

# Base (row 19), Downside (row 31), Upside (row 43)
for row_num, label, color, dash in [
    (19, "Base", NAVY, None),
    (31, "Downside", RED, "dash"),
    (43, "Upside", GREEN, "dash"),
]:
    ref = Reference(ws, min_col=2, max_col=11, min_row=row_num)
    chart1.add_data(ref, from_rows=True)
    idx = len(chart1.series) - 1
    chart1.series[idx].title = openpyxl.chart.series.SeriesLabel(v=label)
    chart1.series[idx].graphicalProperties.line.solidFill = color
    chart1.series[idx].graphicalProperties.line.width = 28000 if not dash else 20000
    if dash:
        chart1.series[idx].graphicalProperties.line.dashStyle = dash
    chart1.series[idx].smooth = True

chart1.set_categories(cats)
chart1.series[0].dLbls = DataLabelList()
chart1.series[0].dLbls.showVal = True
chart1.series[0].dLbls.numFmt = '#,##0'
chart1.series[0].dLbls.showCatName = False
chart1.series[0].dLbls.showSerName = False

# Chart 2: Cumulative Cash (Line)
chart2 = LineChart()
chart2.title = "Cumulative Cash (10-Year)"
chart2.style = 2
chart2.y_axis.title = "CHF"
chart2.y_axis.numFmt = '#,##0'
chart2.height = 12
chart2.width = 20
chart2.legend.position = 'b'
chart2.y_axis.crossesAt = 0

# Base (row 20), Downside (row 32), Upside (row 44)
for row_num, label, color, dash in [
    (20, "Base", NAVY, None),
    (32, "Downside", RED, "dash"),
    (44, "Upside", GREEN, "dash"),
]:
    ref = Reference(ws, min_col=2, max_col=11, min_row=row_num)
    chart2.add_data(ref, from_rows=True)
    idx = len(chart2.series) - 1
    chart2.series[idx].title = openpyxl.chart.series.SeriesLabel(v=label)
    chart2.series[idx].graphicalProperties.line.solidFill = color
    chart2.series[idx].graphicalProperties.line.width = 28000 if not dash else 20000
    if dash:
        chart2.series[idx].graphicalProperties.line.dashStyle = dash
    chart2.series[idx].smooth = True

chart2.set_categories(cats)
chart2.series[0].dLbls = DataLabelList()
chart2.series[0].dLbls.showVal = True
chart2.series[0].dLbls.numFmt = '#,##0'
chart2.series[0].dLbls.showCatName = False
chart2.series[0].dLbls.showSerName = False

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
