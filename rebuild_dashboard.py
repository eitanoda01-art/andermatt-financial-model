import openpyxl
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties, LineEndProperties
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers, Protection
from openpyxl.utils import get_column_letter
from copy import copy
import shutil

SRC = "/root/.claude/uploads/ca8257f0-1581-55bf-88a4-34c1d47ea869/fb5c87e5-Aura_Andermatt_10Year_Model.xlsx"
DST = "/home/user/andermatt-financial-model/Aura_Andermatt_10Year_Model.xlsx"

shutil.copy2(SRC, DST)
wb = openpyxl.load_workbook(DST)

# ── Colour palette ──
NAVY = "1F4E79"
LIGHT_BLUE = "D6E4F0"
WHITE = "FFFFFF"
YELLOW = "FFFF00"
LIGHT_GREY = "F2F2F2"
MID_GREY = "D9D9D9"
GREEN_ACCENT = "2E7D32"
RED_ACCENT = "C62828"
DARK = "333333"

fill_navy = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
fill_light_blue = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
fill_white = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
fill_yellow = PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")
fill_grey = PatternFill(start_color=LIGHT_GREY, end_color=LIGHT_GREY, fill_type="solid")
fill_mid_grey = PatternFill(start_color=MID_GREY, end_color=MID_GREY, fill_type="solid")

font_title = Font(name="Calibri", size=18, bold=True, color=NAVY)
font_section = Font(name="Calibri", size=13, bold=True, color=NAVY)
font_kpi_label = Font(name="Calibri", size=10, color="666666")
font_kpi_value = Font(name="Calibri", size=16, bold=True, color=DARK)
font_kpi_value_neg = Font(name="Calibri", size=16, bold=True, color=RED_ACCENT)
font_header = Font(name="Calibri", size=10, bold=True, color=WHITE)
font_row_label = Font(name="Calibri", size=10, color=DARK)
font_row_bold = Font(name="Calibri", size=10, bold=True, color=DARK)
font_small = Font(name="Calibri", size=9, color="888888")
font_yellow_label = Font(name="Calibri", size=10, bold=True, color=NAVY)

thin_border = Border(
    bottom=Side(style="thin", color=MID_GREY)
)
thick_border_bottom = Border(
    bottom=Side(style="medium", color=NAVY)
)

align_right = Alignment(horizontal="right", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_center = Alignment(horizontal="center", vertical="center")

CHF_FMT = '#,##0'
CHF_FMT_NEG = '#,##0;[Red]-#,##0'
PCT_FMT = '0%'

# ── Delete old Dashboard, create new ──
del wb["Dashboard"]
ws = wb.create_sheet("Dashboard", 0)

# Column widths
ws.column_dimensions['A'].width = 32
for col_letter in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
    ws.column_dimensions[col_letter].width = 14

# ══════════════════════════════════════════════
# ROW 1: TITLE
# ══════════════════════════════════════════════
ws.merge_cells('A1:K1')
ws['A1'] = "AURA ANDERMATT — DASHBOARD"
ws['A1'].font = font_title
ws['A1'].alignment = Alignment(vertical="center")
ws.row_dimensions[1].height = 36

# ══════════════════════════════════════════════
# ROWS 3-5: KPI CARDS (4 cards across)
# ══════════════════════════════════════════════
ws.row_dimensions[2].height = 6  # spacer

# Card positions: A3:B5, D3:E5, G3:H5, J3:K5
kpi_configs = [
    {
        'label_cell': 'A3', 'value_cell': 'A4', 'scenario_cell': 'A5',
        'merge_label': 'A3:B3', 'merge_value': 'A4:B4', 'merge_scenario': 'A5:B5',
        'label': '10-Year Cumulative Cash',
        'base_formula': '=K20',
        'down_formula': '=K32',
        'up_formula': '=K44',
    },
    {
        'label_cell': 'D3', 'value_cell': 'D4', 'scenario_cell': 'D5',
        'merge_label': 'D3:E3', 'merge_value': 'D4:E4', 'merge_scenario': 'D5:E5',
        'label': '2027 Profit After Tax',
        'base_formula': '=B19',
        'down_formula': '=B31',
        'up_formula': '=B43',
    },
    {
        'label_cell': 'G3', 'value_cell': 'G4', 'scenario_cell': 'G5',
        'merge_label': 'G3:H3', 'merge_value': 'G4:H4', 'merge_scenario': 'G5:H5',
        'label': '10-Year Total Revenue',
        'base_formula': '=SUM(B15:K15)',
        'down_formula': '=SUM(B27:K27)',
        'up_formula': '=SUM(B39:K39)',
    },
    {
        'label_cell': 'J3', 'value_cell': 'J4', 'scenario_cell': 'J5',
        'merge_label': 'J3:K3', 'merge_value': 'J4:K4', 'merge_scenario': 'J5:K5',
        'label': 'Break-Even Year (Base)',
        'base_formula': None,  # special
        'down_formula': None,
        'up_formula': None,
    },
]

for cfg in kpi_configs:
    ws.merge_cells(cfg['merge_label'])
    ws.merge_cells(cfg['merge_value'])
    ws.merge_cells(cfg['merge_scenario'])

    # Label row
    c = ws[cfg['label_cell']]
    c.value = cfg['label']
    c.font = font_kpi_label
    c.fill = fill_light_blue
    c.alignment = align_center
    # fill the second cell of merge too
    label_row = c.row
    label_col = c.column
    ws.cell(row=label_row, column=label_col+1).fill = fill_light_blue

    # Value row
    vc = ws[cfg['value_cell']]
    vc.font = font_kpi_value
    vc.alignment = align_center
    vc.fill = fill_white
    ws.cell(row=vc.row, column=vc.column+1).fill = fill_white

    # Scenario row (downside / upside small text)
    sc = ws[cfg['scenario_cell']]
    sc.font = font_small
    sc.alignment = align_center
    sc.fill = fill_white
    ws.cell(row=sc.row, column=sc.column+1).fill = fill_white

# Set KPI values
ws['A4'].value = '=K20'
ws['A4'].number_format = CHF_FMT
ws['A5'].value = 'Down: =TEXT(K32,"#,##0") / Up: =TEXT(K44,"#,##0")'
# Actually use a formula for scenario display
ws['A5'] = '="Down: "&TEXT(K32,"#,##0")&"  |  Up: "&TEXT(K44,"#,##0")'
ws['A5'].font = font_small

ws['D4'].value = '=B19'
ws['D4'].number_format = CHF_FMT_NEG
ws['D5'] = '="Down: "&TEXT(B31,"#,##0")&"  |  Up: "&TEXT(B43,"#,##0")'
ws['D5'].font = font_small

ws['G4'].value = '=SUM(B15:K15)'
ws['G4'].number_format = CHF_FMT
ws['G5'] = '="Down: "&TEXT(SUM(B27:K27),"#,##0")&"  |  Up: "&TEXT(SUM(B39:K39),"#,##0")'
ws['G5'].font = font_small

# Break-even year: find first year where cumulative cash > 0
# Cumulative cash is in row 20 (B20:K20 for Base)
# Years are 2027-2036 in B11:K11
# Use a simple IF chain
be_formula = '=IF(B20>0,2027,IF(C20>0,2028,IF(D20>0,2029,IF(E20>0,2030,IF(F20>0,2031,IF(G20>0,2032,IF(H20>0,2033,IF(I20>0,2034,IF(J20>0,2035,IF(K20>0,2036,"N/A"))))))))))'
ws['J4'].value = be_formula
ws['J4'].font = font_kpi_value
ws['J4'].number_format = '0'
ws['J5'] = '="(Cumulative cash turns positive)"'
ws['J5'].font = font_small

# Add borders around KPI cards
for row_num in [3, 4, 5]:
    for col_num in range(1, 12):
        cell = ws.cell(row=row_num, column=col_num)
        if col_num in [3, 6, 9]:  # gap columns
            cell.fill = fill_white

ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 30
ws.row_dimensions[5].height = 18

# ══════════════════════════════════════════════
# ROW 6: spacer
# ══════════════════════════════════════════════
ws.row_dimensions[6].height = 6

# ══════════════════════════════════════════════
# ROWS 7-8: SCENARIO COMPARISON heading + year headers
# ══════════════════════════════════════════════
# We'll build the data tables FIRST (rows 7-52ish), then place charts AFTER the tables

# ── BASE SCENARIO TABLE (rows 7-20) ──
r = 7
ws.merge_cells(f'A{r}:K{r}')
ws[f'A{r}'] = "BASE SCENARIO"
ws[f'A{r}'].font = font_section
ws[f'A{r}'].fill = fill_light_blue
for c in range(2, 12):
    ws.cell(row=r, column=c).fill = fill_light_blue
ws.row_dimensions[r].height = 22

r = 8  # year headers
years = list(range(2027, 2037))
for i, yr in enumerate(years):
    c = ws.cell(row=r, column=i+2)
    c.value = yr
    c.font = font_header
    c.fill = fill_navy
    c.alignment = align_center
ws.cell(row=r, column=1).fill = fill_navy
ws.cell(row=r, column=1).font = font_header

# Revenue lines (rows 9-12)
rev_rows = [
    (9, '  Net Ticket Revenue', "='P&L'!{col}10*$C$47"),
    (10, '  Partnership Revenue', "='P&L'!{col}11*$C$48"),
    (11, '  Other Revenue (ASA+F&B+Other)', "='P&L'!{col}12+'P&L'!{col}13+'P&L'!{col}14+'P&L'!{col}15"),
    (12, 'Total Revenue', None),  # special formula with summer
]
for row_num, label, formula_tmpl in rev_rows:
    ws.cell(row=row_num, column=1).value = label
    if row_num == 12:
        ws.cell(row=row_num, column=1).font = font_row_bold
    else:
        ws.cell(row=row_num, column=1).font = font_row_label
    for i, yr in enumerate(years):
        col_letter = get_column_letter(i+2)
        cell = ws.cell(row=row_num, column=i+2)
        if row_num == 12:
            cell.value = f'=({col_letter}9+{col_letter}10+{col_letter}11)*(1+IF({yr}>=Revenue!$B$13,Revenue!$B$11,0))'
            cell.font = font_row_bold
            cell.border = thick_border_bottom
        else:
            cell.value = formula_tmpl.format(col=col_letter)
            cell.font = font_row_label
            cell.border = thin_border
        cell.number_format = CHF_FMT_NEG
        cell.alignment = align_right

# Cost, EBIT, Tax, Profit After Tax, Cumulative Cash
result_rows = [
    (13, '  Total Costs', "='P&L'!{col}24*$C$49", font_row_label),
    (14, 'EBIT', "={col}12+{col}13", font_row_bold),
    (15, '  Tax', "=IF({col}14>0,-{col}14*'P&L'!$B$5,0)", font_row_label),
    (16, 'Profit After Tax', "={col}14+{col}15", font_row_bold),
    (17, 'Cumulative Cash', None, font_row_bold),
]
for row_num, label, formula_tmpl, font in result_rows:
    ws.cell(row=row_num, column=1).value = label
    ws.cell(row=row_num, column=1).font = font
    for i, yr in enumerate(years):
        col_letter = get_column_letter(i+2)
        cell = ws.cell(row=row_num, column=i+2)
        if row_num == 17:
            if i == 0:
                cell.value = f"='P&L'!$B$6+{col_letter}16"
            else:
                prev = get_column_letter(i+1)
                cell.value = f"={prev}17+{col_letter}16"
        else:
            cell.value = formula_tmpl.format(col=col_letter)
        cell.font = font
        cell.number_format = CHF_FMT_NEG
        cell.alignment = align_right
        if row_num in [14, 16, 17]:
            cell.border = thick_border_bottom

# ── DOWNSIDE SCENARIO (rows 19-29) ──
r = 19
ws.row_dimensions[18].height = 10  # spacer
ws.merge_cells(f'A{r}:K{r}')
ws[f'A{r}'] = "DOWNSIDE SCENARIO"
ws[f'A{r}'].font = font_section
ws[f'A{r}'].fill = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
for c in range(2, 12):
    ws.cell(row=r, column=c).fill = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
ws.row_dimensions[r].height = 22

r = 20
for i, yr in enumerate(years):
    c = ws.cell(row=r, column=i+2)
    c.value = yr
    c.font = font_header
    c.fill = PatternFill(start_color=RED_ACCENT, end_color=RED_ACCENT, fill_type="solid")
    c.alignment = align_center
ws.cell(row=r, column=1).fill = PatternFill(start_color=RED_ACCENT, end_color=RED_ACCENT, fill_type="solid")

down_rows = [
    (21, '  Net Ticket Revenue (adj.)', "='P&L'!{col}10*$B$47"),
    (22, '  Partnership Revenue (adj.)', "='P&L'!{col}11*$B$48"),
    (23, '  Other Revenue (ASA+F&B+Other)', "='P&L'!{col}12+'P&L'!{col}13+'P&L'!{col}14+'P&L'!{col}15"),
    (24, 'Total Revenue', None),
    (25, '  Total Costs (adj.)', "='P&L'!{col}24*$B$49", font_row_label),
    (26, 'EBIT', "={col}24+{col}25", font_row_bold),
    (27, '  Tax', "=IF({col}26>0,-{col}26*'P&L'!$B$5,0)", font_row_label),
    (28, 'Profit After Tax', "={col}26+{col}27", font_row_bold),
    (29, 'Cumulative Cash', None, font_row_bold),
]
for item in down_rows:
    row_num = item[0]
    label = item[1]
    formula_tmpl = item[2]
    font = item[3] if len(item) > 3 else (font_row_bold if 'Total' in label or 'EBIT' in label or 'Profit' in label or 'Cumulative' in label else font_row_label)
    ws.cell(row=row_num, column=1).value = label
    ws.cell(row=row_num, column=1).font = font
    for i, yr in enumerate(years):
        col_letter = get_column_letter(i+2)
        cell = ws.cell(row=row_num, column=i+2)
        if row_num == 24:
            cell.value = f'=({col_letter}21+{col_letter}22+{col_letter}23)*(1+IF({yr}>=Revenue!$B$13,Revenue!$B$11,0))'
            cell.font = font_row_bold
            cell.border = thick_border_bottom
        elif row_num == 29:
            if i == 0:
                cell.value = f"='P&L'!$B$6+{col_letter}28"
            else:
                prev = get_column_letter(i+1)
                cell.value = f"={prev}29+{col_letter}28"
            cell.font = font_row_bold
        else:
            cell.value = formula_tmpl.format(col=col_letter)
            cell.font = font
        cell.number_format = CHF_FMT_NEG
        cell.alignment = align_right
        if row_num in [26, 28, 29]:
            cell.border = thick_border_bottom

# ── UPSIDE SCENARIO (rows 31-41) ──
r = 31
ws.row_dimensions[30].height = 10
ws.merge_cells(f'A{r}:K{r}')
ws[f'A{r}'] = "UPSIDE SCENARIO"
ws[f'A{r}'].font = font_section
ws[f'A{r}'].fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
for c in range(2, 12):
    ws.cell(row=r, column=c).fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
ws.row_dimensions[r].height = 22

r = 32
for i, yr in enumerate(years):
    c = ws.cell(row=r, column=i+2)
    c.value = yr
    c.font = font_header
    c.fill = PatternFill(start_color=GREEN_ACCENT, end_color=GREEN_ACCENT, fill_type="solid")
    c.alignment = align_center
ws.cell(row=r, column=1).fill = PatternFill(start_color=GREEN_ACCENT, end_color=GREEN_ACCENT, fill_type="solid")

up_rows = [
    (33, '  Net Ticket Revenue (adj.)', "='P&L'!{col}10*$D$47"),
    (34, '  Partnership Revenue (adj.)', "='P&L'!{col}11*$D$48"),
    (35, '  Other Revenue (ASA+F&B+Other)', "='P&L'!{col}12+'P&L'!{col}13+'P&L'!{col}14+'P&L'!{col}15"),
    (36, 'Total Revenue', None),
    (37, '  Total Costs (adj.)', "='P&L'!{col}24*$D$49", font_row_label),
    (38, 'EBIT', "={col}36+{col}37", font_row_bold),
    (39, '  Tax', "=IF({col}38>0,-{col}38*'P&L'!$B$5,0)", font_row_label),
    (40, 'Profit After Tax', "={col}38+{col}39", font_row_bold),
    (41, 'Cumulative Cash', None, font_row_bold),
]
for item in up_rows:
    row_num = item[0]
    label = item[1]
    formula_tmpl = item[2]
    font = item[3] if len(item) > 3 else (font_row_bold if 'Total' in label or 'EBIT' in label or 'Profit' in label or 'Cumulative' in label else font_row_label)
    ws.cell(row=row_num, column=1).value = label
    ws.cell(row=row_num, column=1).font = font
    for i, yr in enumerate(years):
        col_letter = get_column_letter(i+2)
        cell = ws.cell(row=row_num, column=i+2)
        if row_num == 36:
            cell.value = f'=({col_letter}33+{col_letter}34+{col_letter}35)*(1+IF({yr}>=Revenue!$B$13,Revenue!$B$11,0))'
            cell.font = font_row_bold
            cell.border = thick_border_bottom
        elif row_num == 41:
            if i == 0:
                cell.value = f"='P&L'!$B$6+{col_letter}40"
            else:
                prev = get_column_letter(i+1)
                cell.value = f"={prev}41+{col_letter}40"
            cell.font = font_row_bold
        else:
            cell.value = formula_tmpl.format(col=col_letter)
            cell.font = font
        cell.number_format = CHF_FMT_NEG
        cell.alignment = align_right
        if row_num in [38, 40, 41]:
            cell.border = thick_border_bottom

# ══════════════════════════════════════════════
# SCENARIO ADJUSTMENTS (rows 44-55)
# ══════════════════════════════════════════════
ws.row_dimensions[42].height = 6
ws.row_dimensions[43].height = 10

r = 44
ws.merge_cells(f'A{r}:K{r}')
ws[f'A{r}'] = "SCENARIO ADJUSTMENTS"
ws[f'A{r}'].font = font_section
ws[f'A{r}'].fill = fill_light_blue
for c in range(2, 12):
    ws.cell(row=r, column=c).fill = fill_light_blue
ws.row_dimensions[r].height = 22

# Headers
r = 45
for col_num, (label, fill) in [(2, ('Downside', PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid"))),
                                  (3, ('Base', fill_light_blue)),
                                  (4, ('Upside', PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")))]:
    c = ws.cell(row=r, column=col_num)
    c.value = label
    c.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
    c.fill = fill
    c.alignment = align_center

ws.cell(row=r, column=1).fill = fill_grey

# Adjustment rows with explanations
adj_items = [
    (47, 'Ticket Occupancy', 0.8, 1.0, 1.0,
     '% of tickets sold (e.g. 80% = 200 of 250)'),
    (48, 'Partnership Factor', 0.6, 1.0, 1.2,
     'Multiplier on partnership revenue'),
    (49, 'Cost Overrun Factor', 1.15, 1.0, 0.95,
     'Multiplier on total costs (1.15 = 15% over budget)'),
]

r = 46
ws.cell(row=r, column=1).value = "Adjust these values to change scenarios:"
ws.cell(row=r, column=1).font = font_small

for row_num, label, down_val, base_val, up_val, explanation in adj_items:
    ws.cell(row=row_num, column=1).value = label
    ws.cell(row=row_num, column=1).font = font_yellow_label

    for col_num, val in [(2, down_val), (3, base_val), (4, up_val)]:
        c = ws.cell(row=row_num, column=col_num)
        c.value = val
        c.number_format = '0.00'
        c.alignment = align_center
        c.font = Font(name="Calibri", size=11, bold=True)
        c.fill = fill_yellow

    # Explanation in column E onwards
    ws.merge_cells(f'E{row_num}:K{row_num}')
    ws.cell(row=row_num, column=5).value = explanation
    ws.cell(row=row_num, column=5).font = font_small

# Note
r = 51
ws.merge_cells(f'A{r}:K{r}')
ws[f'A{r}'] = "Downside = worst case. Base = as entered in Revenue/Cost sheets. Upside = best case. To change prices/growth rates, edit yellow cells in Revenue/Cost sheets."
ws[f'A{r}'].font = font_small

# ══════════════════════════════════════════════
# CHARTS — placed ABOVE the tables, in columns L-T
# Actually, let's put them at the right side of the KPI area
# Better: put 3 charts side by side spanning rows 7 area
# ══════════════════════════════════════════════

# Chart 1: Total Revenue (Line) — 3 scenarios
chart1 = LineChart()
chart1.title = "Total Revenue"
chart1.style = 2
chart1.y_axis.title = "CHF"
chart1.y_axis.numFmt = '#,##0'
chart1.x_axis.title = None
chart1.height = 14
chart1.width = 22
chart1.legend.position = 'b'

cats = Reference(ws, min_col=2, max_col=11, min_row=8)  # years

s1 = Reference(ws, min_col=2, max_col=11, min_row=12)  # Base Total Revenue
chart1.add_data(s1, from_rows=True)
chart1.series[0].title = openpyxl.chart.series.SeriesLabel(v="Base")
chart1.series[0].graphicalProperties.line.solidFill = NAVY
chart1.series[0].graphicalProperties.line.width = 28000
chart1.series[0].dLbls = DataLabelList()
chart1.series[0].dLbls.showVal = False

s2 = Reference(ws, min_col=2, max_col=11, min_row=24)  # Downside Total Revenue
chart1.add_data(s2, from_rows=True)
chart1.series[1].title = openpyxl.chart.series.SeriesLabel(v="Downside")
chart1.series[1].graphicalProperties.line.solidFill = RED_ACCENT
chart1.series[1].graphicalProperties.line.dashStyle = "dash"
chart1.series[1].graphicalProperties.line.width = 20000

s3 = Reference(ws, min_col=2, max_col=11, min_row=36)  # Upside Total Revenue
chart1.add_data(s3, from_rows=True)
chart1.series[2].title = openpyxl.chart.series.SeriesLabel(v="Upside")
chart1.series[2].graphicalProperties.line.solidFill = GREEN_ACCENT
chart1.series[2].graphicalProperties.line.dashStyle = "dash"
chart1.series[2].graphicalProperties.line.width = 20000

chart1.set_categories(cats)

# Show data labels on Base only (first and last point)
chart1.series[0].dLbls = DataLabelList()
chart1.series[0].dLbls.showVal = True
chart1.series[0].dLbls.numFmt = '#,##0'
chart1.series[0].dLbls.showCatName = False
chart1.series[0].dLbls.showSerName = False

# Chart 2: EBIT (Bar chart)
chart2 = BarChart()
chart2.title = "EBIT"
chart2.style = 2
chart2.y_axis.title = "CHF"
chart2.y_axis.numFmt = '#,##0'
chart2.height = 14
chart2.width = 22
chart2.legend.position = 'b'

cats2 = Reference(ws, min_col=2, max_col=11, min_row=8)

s1 = Reference(ws, min_col=2, max_col=11, min_row=14)  # Base EBIT
chart2.add_data(s1, from_rows=True)
chart2.series[0].title = openpyxl.chart.series.SeriesLabel(v="Base")
chart2.series[0].graphicalProperties.solidFill = NAVY

s2 = Reference(ws, min_col=2, max_col=11, min_row=26)  # Downside EBIT
chart2.add_data(s2, from_rows=True)
chart2.series[1].title = openpyxl.chart.series.SeriesLabel(v="Downside")
chart2.series[1].graphicalProperties.solidFill = RED_ACCENT

s3 = Reference(ws, min_col=2, max_col=11, min_row=38)  # Upside EBIT
chart2.add_data(s3, from_rows=True)
chart2.series[2].title = openpyxl.chart.series.SeriesLabel(v="Upside")
chart2.series[2].graphicalProperties.solidFill = GREEN_ACCENT

chart2.set_categories(cats2)

# Add zero reference line
chart2.y_axis.crossesAt = 0

# Data labels on Base
chart2.series[0].dLbls = DataLabelList()
chart2.series[0].dLbls.showVal = True
chart2.series[0].dLbls.numFmt = '#,##0'
chart2.series[0].dLbls.showCatName = False
chart2.series[0].dLbls.showSerName = False

# Chart 3: Cumulative Cash (Line)
chart3 = LineChart()
chart3.title = "Cumulative Cash"
chart3.style = 2
chart3.y_axis.title = "CHF"
chart3.y_axis.numFmt = '#,##0'
chart3.height = 14
chart3.width = 22
chart3.legend.position = 'b'

cats3 = Reference(ws, min_col=2, max_col=11, min_row=8)

s1 = Reference(ws, min_col=2, max_col=11, min_row=17)  # Base Cum Cash
chart3.add_data(s1, from_rows=True)
chart3.series[0].title = openpyxl.chart.series.SeriesLabel(v="Base")
chart3.series[0].graphicalProperties.line.solidFill = NAVY
chart3.series[0].graphicalProperties.line.width = 28000

s2 = Reference(ws, min_col=2, max_col=11, min_row=29)  # Downside
chart3.add_data(s2, from_rows=True)
chart3.series[1].title = openpyxl.chart.series.SeriesLabel(v="Downside")
chart3.series[1].graphicalProperties.line.solidFill = RED_ACCENT
chart3.series[1].graphicalProperties.line.dashStyle = "dash"
chart3.series[1].graphicalProperties.line.width = 20000

s3 = Reference(ws, min_col=2, max_col=11, min_row=41)  # Upside
chart3.add_data(s3, from_rows=True)
chart3.series[2].title = openpyxl.chart.series.SeriesLabel(v="Upside")
chart3.series[2].graphicalProperties.line.solidFill = GREEN_ACCENT
chart3.series[2].graphicalProperties.line.dashStyle = "dash"
chart3.series[2].graphicalProperties.line.width = 20000

chart3.set_categories(cats3)

# Data labels on Base
chart3.series[0].dLbls = DataLabelList()
chart3.series[0].dLbls.showVal = True
chart3.series[0].dLbls.numFmt = '#,##0'
chart3.series[0].dLbls.showCatName = False
chart3.series[0].dLbls.showSerName = False

# Place charts to the right of the data (columns L onwards)
# Chart 1 top-right
ws.add_chart(chart1, "L1")
# Chart 2 middle-right
ws.add_chart(chart2, "L16")
# Chart 3 bottom-right
ws.add_chart(chart3, "L31")

# ══════════════════════════════════════════════
# P&L SHEET: Rename "NET RESULT" to "PROFIT AFTER TAX"
# ══════════════════════════════════════════════
ws_pl = wb["P&L"]
ws_pl['A29'].value = "PROFIT AFTER TAX"

# ══════════════════════════════════════════════
# SHEET PROTECTION
# ══════════════════════════════════════════════
PASSWORD = "2027"

# Helper: set all cells to locked, then unlock yellow cells
for sheet_name in ["Dashboard", "P&L", "Revenue", "Cost"]:
    sheet = wb[sheet_name]

    # First, set all cells to locked (default)
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell in row:
            cell.protection = Protection(locked=True)

    # Then unlock yellow cells
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell in row:
            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                rgb = str(cell.fill.start_color.rgb)
                if rgb in ("FFFFFF00", "00FFFF00"):
                    cell.protection = Protection(locked=False)

    # Enable sheet protection
    sheet.protection.sheet = True
    sheet.protection.password = PASSWORD
    sheet.protection.enable()
    # Allow selecting cells
    sheet.protection.selectLockedCells = False
    sheet.protection.selectUnlockedCells = False

# Dashboard scenario adjustment cells are yellow — already unlocked above
# But let's make sure they're explicitly unlocked
for r in [47, 48, 49]:
    for c in [2, 3, 4]:
        ws.cell(row=r, column=c).protection = Protection(locked=False)

# ══════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════
wb.save(DST)
print("Done! Saved to", DST)
