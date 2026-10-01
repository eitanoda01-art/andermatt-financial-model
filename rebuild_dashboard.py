import openpyxl
from openpyxl.styles import Protection
import shutil

SRC = "/root/.claude/uploads/ca8257f0-1581-55bf-88a4-34c1d47ea869/fb5c87e5-Aura_Andermatt_10Year_Model.xlsx"
DST = "/home/user/andermatt-financial-model/Aura_Andermatt_10Year_Model.xlsx"

shutil.copy2(SRC, DST)
wb = openpyxl.load_workbook(DST)

# ── Rename "Net Result" → "Profit After Tax" ──
ws = wb["Dashboard"]
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
    for cell in row:
        if isinstance(cell.value, str) and cell.value.strip() in ("Net Result", "2027 Net Result (CHF)"):
            cell.value = cell.value.replace("Net Result", "Profit After Tax")

ws_pl = wb["P&L"]
if ws_pl["A29"].value and "NET RESULT" in str(ws_pl["A29"].value).upper():
    ws_pl["A29"].value = "PROFIT AFTER TAX"

# ── Sheet protection (password: 2027) ──
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
