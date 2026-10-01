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
GOLD = "D4A017"

fill_navy = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
fill_light_blue = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
fill_white = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
fill_yellow = PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")
fill_grey = PatternFill(start_color=LIGHT_GREY, end_color=LIGHT_GREY, fill_type="solid")
fill_mid_grey = PatternFill(start_color=MID_GREY, end_color=MID_GREY, fill_type="solid")
fill_down = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
fill_up = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

font_title = Font(name="Calibri", size=18, bold=True, color=WHITE)
font_section = Font(name="Calibri", size=12, bold=True, color=NAVY)
font_kpi_label = Font(name="Calibri", size=9, color="666666")
font_kpi_value = Font(name="Calibri", size=18, bold=True, color=NAVY)
font_kpi_sub = Font(name="Calibri", size=8, color="888888")
font_header = Font(name="Calibri", size=10, bold=True, color=WHITE)
font_row_label = Font(name="Calibri", size=10, color=DARK)
font_row_bold = Font(name="Calibri", size=10, bold=True, color=DARK)
font_small = Font(name="Calibri", size=9, color="888888")
font_yellow_label = Font(name="Calibri", size=10, bold=True, color=NAVY)

thin_border = Border(bottom=Side(style="thin", color=MID_GREY))
thick_border_bottom = Border(bottom=Side(style="medium", color=NAVY))

kpi_box_border = Border(
    top=Side(style="thin", color=NAVY),
    bottom=Side(style="thin", color=NAVY),
    left=Side(style="thin", color=NAVY),
    right=Side(style="thin", color=NAVY),
)
kpi_box_border_left = Border(
    top=Side(style="thin", color=NAVY),
    bottom=Side(style="thin", color=NAVY),
    left=Side(style="thin", color=NAVY),
)
kpi_box_border_right = Border(
    top=Side(style="thin", color=NAVY),
    bottom=Side(style="thin", color=NAVY),
    right=Side(style="thin", color=NAVY),
)
kpi_box_border_mid = Border(
    top=Side(style="thin", color=NAVY),
    bottom=Side(style="thin", color=NAVY),
)

align_right = Alignment(horizontal="right", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_center = Alignment(horizontal="center", vertical="center")

CHF_FMT = '#,##0'
CHF_FMT_NEG = '#,##0;[Red]-#,##0'
PCT_FMT = '0%'

years = list(range(2027, 2037))

# ══════════════════════════════════════════════════════════════════
# ROW MAP (where everything lives)
# ══════════════════════════════════════════════════════════════════
# Row 1     : Title bar
# Row 2     : spacer
# Rows 3-5  : KPI cards
# Row 6     : spacer
# Rows 7-37 : Chart display area (empty rows, 3 floating charts)
# Row 38    : spacer
# Rows 39-49: BASE SCENARIO table
# Row 50    : spacer
# Rows 51-61: DOWNSIDE SCENARIO table
# Row 62    : spacer
# Rows 63-73: UPSIDE SCENARIO table
# Row 74    : spacer
# Rows 75-82: SCENARIO ADJUSTMENTS
#
# BASE rows:
#   39=header, 40=years, 41=ticket, 42=partner, 43=other,
#   44=total_rev, 45=costs, 46=EBIT, 47=tax, 48=PAT, 49=cum_cash
# DOWN rows: +12 offset → 51=header, 52=years, 53..61
# UP rows:   +24 offset → 63=header, 64=years, 65..73
# ADJ rows:  75=header, 76=col_headers, 77=instruction,
#            78=ticket_occ, 79=partner_fac, 80=cost_overrun, 82=note
# ══════════════════════════════════════════════════════════════════

# Row numbers for data
B_HDR, B_YR = 39, 40
B_TICK, B_PART, B_OTH, B_REV = 41, 42, 43, 44
B_COST, B_EBIT, B_TAX, B_PAT, B_CUM = 45, 46, 47, 48, 49

D_HDR, D_YR = 51, 52
D_TICK, D_PART, D_OTH, D_REV = 53, 54, 55, 56
D_COST, D_EBIT, D_TAX, D_PAT, D_CUM = 57, 58, 59, 60, 61

U_HDR, U_YR = 63, 64
U_TICK, U_PART, U_OTH, U_REV = 65, 66, 67, 68
U_COST, U_EBIT, U_TAX, U_PAT, U_CUM = 69, 70, 71, 72, 73

ADJ_HDR = 75
ADJ_COL = 76
ADJ_INST = 77
ADJ_TICK, ADJ_PART, ADJ_COST = 78, 79, 80
ADJ_NOTE = 82

# ── Delete old Dashboard, create new ──
del wb["Dashboard"]
ws = wb.create_sheet("Dashboard", 0)

# Column widths
ws.column_dimensions['A'].width = 30
for col_letter in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
    ws.column_dimensions[col_letter].width = 14

# ══════════════════════════════════════════════
# ROW 1: TITLE BAR (navy background)
# ══════════════════════════════════════════════
ws.merge_cells('A1:K1')
ws['A1'] = "AURA ANDERMATT — 10-Year Financial Dashboard"
ws['A1'].font = font_title
ws['A1'].fill = fill_navy
ws['A1'].alignment = Alignment(horizontal="left", vertical="center", indent=1)
for c in range(2, 12):
    ws.cell(row=1, column=c).fill = fill_navy
ws.row_dimensions[1].height = 40

# ══════════════════════════════════════════════
# ROW 2: spacer
# ══════════════════════════════════════════════
ws.row_dimensions[2].height = 8

# ══════════════════════════════════════════════
# ROWS 3-5: KPI CARDS with proper box borders
# ══════════════════════════════════════════════
# 4 cards: A3:B5, D3:E5, G3:H5, J3:K5
# Gap columns: C, F, I (width=2)
for gap_col in ['C', 'F', 'I']:
    ws.column_dimensions[gap_col].width = 2

ws.row_dimensions[3].height = 18
ws.row_dimensions[4].height = 32
ws.row_dimensions[5].height = 16

kpi_defs = [
    {
        'cols': (1, 2), 'merge_l': 'A3:B3', 'merge_v': 'A4:B4', 'merge_s': 'A5:B5',
        'label': 'CUMULATIVE CASH (10Y)',
        'formula': f'=K{B_CUM}',
        'sub': f'="Down: "&TEXT(K{D_CUM},"#,##0")&"  |  Up: "&TEXT(K{U_CUM},"#,##0")',
        'fmt': CHF_FMT_NEG,
    },
    {
        'cols': (4, 5), 'merge_l': 'D3:E3', 'merge_v': 'D4:E4', 'merge_s': 'D5:E5',
        'label': '2027 PROFIT AFTER TAX',
        'formula': f'=B{B_PAT}',
        'sub': f'="Down: "&TEXT(B{D_PAT},"#,##0")&"  |  Up: "&TEXT(B{U_PAT},"#,##0")',
        'fmt': CHF_FMT_NEG,
    },
    {
        'cols': (7, 8), 'merge_l': 'G3:H3', 'merge_v': 'G4:H4', 'merge_s': 'G5:H5',
        'label': 'TOTAL REVENUE (10Y)',
        'formula': f'=SUM(B{B_REV}:K{B_REV})',
        'sub': f'="Down: "&TEXT(SUM(B{D_REV}:K{D_REV}),"#,##0")&"  |  Up: "&TEXT(SUM(B{U_REV}:K{U_REV}),"#,##0")',
        'fmt': CHF_FMT,
    },
    {
        'cols': (10, 11), 'merge_l': 'J3:K3', 'merge_v': 'J4:K4', 'merge_s': 'J5:K5',
        'label': 'BREAK-EVEN YEAR',
        'formula': f'=IF(B{B_CUM}>0,2027,IF(C{B_CUM}>0,2028,IF(D{B_CUM}>0,2029,IF(E{B_CUM}>0,2030,IF(F{B_CUM}>0,2031,IF(G{B_CUM}>0,2032,IF(H{B_CUM}>0,2033,IF(I{B_CUM}>0,2034,IF(J{B_CUM}>0,2035,IF(K{B_CUM}>0,2036,"N/A"))))))))))',
        'sub': '="(Cumulative cash turns positive)"',
        'fmt': '0',
    },
]

fill_kpi_bg = PatternFill(start_color="F7F9FC", end_color="F7F9FC", fill_type="solid")
fill_kpi_top = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")

for kpi in kpi_defs:
    c1, c2 = kpi['cols']
    ws.merge_cells(kpi['merge_l'])
    ws.merge_cells(kpi['merge_v'])
    ws.merge_cells(kpi['merge_s'])

    # Label row (navy strip at top of card)
    lbl = ws.cell(row=3, column=c1)
    lbl.value = kpi['label']
    lbl.font = Font(name="Calibri", size=8, bold=True, color=WHITE)
    lbl.fill = fill_kpi_top
    lbl.alignment = align_center
    ws.cell(row=3, column=c2).fill = fill_kpi_top

    # Value row
    val = ws.cell(row=4, column=c1)
    val.value = kpi['formula']
    val.font = font_kpi_value
    val.alignment = align_center
    val.fill = fill_kpi_bg
    val.number_format = kpi['fmt']
    ws.cell(row=4, column=c2).fill = fill_kpi_bg

    # Sub row
    sub = ws.cell(row=5, column=c1)
    sub.value = kpi['sub']
    sub.font = font_kpi_sub
    sub.alignment = align_center
    sub.fill = fill_kpi_bg
    ws.cell(row=5, column=c2).fill = fill_kpi_bg

    # Box borders on all 3 rows
    for row in [3, 4, 5]:
        ws.cell(row=row, column=c1).border = Border(
            top=Side(style="thin", color=NAVY) if row == 3 else Side(style=None),
            bottom=Side(style="thin", color=NAVY) if row == 5 else Side(style=None),
            left=Side(style="thin", color=NAVY),
        )
        ws.cell(row=row, column=c2).border = Border(
            top=Side(style="thin", color=NAVY) if row == 3 else Side(style=None),
            bottom=Side(style="thin", color=NAVY) if row == 5 else Side(style=None),
            right=Side(style="thin", color=NAVY),
        )

# ══════════════════════════════════════════════
# ROW 6: spacer before charts
# ══════════════════════════════════════════════
ws.row_dimensions[6].height = 15

# Rows 7-37: leave empty for chart display
for r in range(7, 38):
    ws.row_dimensions[r].height = 15

ws.row_dimensions[38].height = 8  # spacer before data

# ══════════════════════════════════════════════
# HELPER: build a scenario data table
# ══════════════════════════════════════════════
def build_scenario_table(ws, hdr_row, name, color_fill, color_header_fill,
                         tick_adj_col, part_adj_col, cost_adj_col):
    yr_row = hdr_row + 1
    tick_row = hdr_row + 2
    part_row = hdr_row + 3
    oth_row = hdr_row + 4
    rev_row = hdr_row + 5
    cost_row = hdr_row + 6
    ebit_row = hdr_row + 7
    tax_row = hdr_row + 8
    pat_row = hdr_row + 9
    cum_row = hdr_row + 10

    # Section header
    ws.merge_cells(f'A{hdr_row}:K{hdr_row}')
    ws[f'A{hdr_row}'] = name
    ws[f'A{hdr_row}'].font = font_section
    ws[f'A{hdr_row}'].fill = color_fill
    for c in range(2, 12):
        ws.cell(row=hdr_row, column=c).fill = color_fill
    ws.row_dimensions[hdr_row].height = 24

    # Year headers
    ws.cell(row=yr_row, column=1).fill = color_header_fill
    ws.cell(row=yr_row, column=1).font = font_header
    for i, yr in enumerate(years):
        c = ws.cell(row=yr_row, column=i+2)
        c.value = yr
        c.font = font_header
        c.fill = color_header_fill
        c.alignment = align_center

    # Data rows
    row_defs = [
        (tick_row, 'Net Ticket Revenue', f"='P&L'!{{col}}10*${tick_adj_col}${ADJ_TICK}", False),
        (part_row, 'Partnership Revenue', f"='P&L'!{{col}}11*${part_adj_col}${ADJ_PART}", False),
        (oth_row, 'Other Revenue', "='P&L'!{col}12+'P&L'!{col}13+'P&L'!{col}14+'P&L'!{col}15", False),
        (rev_row, 'Total Revenue', None, True),
        (cost_row, 'Total Costs', f"='P&L'!{{col}}24*${cost_adj_col}${ADJ_COST}", False),
        (ebit_row, 'EBIT', "={col}{rev}+{col}{cost}", True),
        (tax_row, 'Tax', "=IF({col}{ebit}>0,-{col}{ebit}*'P&L'!$B$5,0)", False),
        (pat_row, 'Profit After Tax', "={col}{ebit}+{col}{tax}", True),
        (cum_row, 'Cumulative Cash', None, True),
    ]

    for rn, label, tmpl, bold in row_defs:
        ws.cell(row=rn, column=1).value = ('  ' if not bold else '') + label
        ws.cell(row=rn, column=1).font = font_row_bold if bold else font_row_label

        for i, yr in enumerate(years):
            col = get_column_letter(i + 2)
            cell = ws.cell(row=rn, column=i + 2)

            if rn == rev_row:
                cell.value = f'=({col}{tick_row}+{col}{part_row}+{col}{oth_row})*(1+IF({yr}>=Revenue!$B$13,Revenue!$B$11,0))'
            elif rn == ebit_row:
                cell.value = f'={col}{rev_row}+{col}{cost_row}'
            elif rn == tax_row:
                cell.value = f"=IF({col}{ebit_row}>0,-{col}{ebit_row}*'P&L'!$B$5,0)"
            elif rn == pat_row:
                cell.value = f'={col}{ebit_row}+{col}{tax_row}'
            elif rn == cum_row:
                if i == 0:
                    cell.value = f"='P&L'!$B$6+{col}{pat_row}"
                else:
                    prev = get_column_letter(i + 1)
                    cell.value = f"={prev}{cum_row}+{col}{pat_row}"
            else:
                cell.value = tmpl.format(col=col)

            cell.font = font_row_bold if bold else font_row_label
            cell.number_format = CHF_FMT_NEG
            cell.alignment = align_right

            if bold:
                cell.border = thick_border_bottom
            else:
                cell.border = thin_border

    return {
        'rev': rev_row, 'ebit': ebit_row, 'pat': pat_row, 'cum': cum_row,
        'yr': yr_row,
    }


# ── BUILD ALL THREE SCENARIO TABLES ──
base = build_scenario_table(ws, B_HDR, "BASE SCENARIO",
                            fill_light_blue, fill_navy, 'C', 'C', 'C')

ws.row_dimensions[50].height = 10
down = build_scenario_table(ws, D_HDR, "DOWNSIDE SCENARIO",
                            fill_down,
                            PatternFill(start_color=RED_ACCENT, end_color=RED_ACCENT, fill_type="solid"),
                            'B', 'B', 'B')

ws.row_dimensions[62].height = 10
up = build_scenario_table(ws, U_HDR, "UPSIDE SCENARIO",
                          fill_up,
                          PatternFill(start_color=GREEN_ACCENT, end_color=GREEN_ACCENT, fill_type="solid"),
                          'D', 'D', 'D')

# ══════════════════════════════════════════════
# SCENARIO ADJUSTMENTS
# ══════════════════════════════════════════════
ws.row_dimensions[74].height = 15

ws.merge_cells(f'A{ADJ_HDR}:K{ADJ_HDR}')
ws[f'A{ADJ_HDR}'] = "SCENARIO ADJUSTMENTS"
ws[f'A{ADJ_HDR}'].font = font_section
ws[f'A{ADJ_HDR}'].fill = fill_light_blue
for c in range(2, 12):
    ws.cell(row=ADJ_HDR, column=c).fill = fill_light_blue
ws.row_dimensions[ADJ_HDR].height = 24

# Column headers
for col_num, (label, fill) in [(2, ('Downside', fill_down)),
                                (3, ('Base', fill_light_blue)),
                                (4, ('Upside', fill_up))]:
    c = ws.cell(row=ADJ_COL, column=col_num)
    c.value = label
    c.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
    c.fill = fill
    c.alignment = align_center
ws.cell(row=ADJ_COL, column=1).fill = fill_grey

ws.cell(row=ADJ_INST, column=1).value = "Adjust yellow cells to change scenarios:"
ws.cell(row=ADJ_INST, column=1).font = font_small

adj_items = [
    (ADJ_TICK, 'Ticket Occupancy', 0.8, 1.0, 1.0,
     '% of tickets sold (e.g. 0.80 = 200 of 250)'),
    (ADJ_PART, 'Partnership Factor', 0.6, 1.0, 1.2,
     'Multiplier on partnership revenue'),
    (ADJ_COST, 'Cost Overrun Factor', 1.15, 1.0, 0.95,
     'Multiplier on total costs (1.15 = +15%)'),
]

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

    ws.merge_cells(f'E{row_num}:K{row_num}')
    ws.cell(row=row_num, column=5).value = explanation
    ws.cell(row=row_num, column=5).font = font_small

ws.merge_cells(f'A{ADJ_NOTE}:K{ADJ_NOTE}')
ws[f'A{ADJ_NOTE}'] = "Downside = worst case. Base = as planned. Upside = best case. To change base prices/growth, edit yellow cells in Revenue/Cost sheets."
ws[f'A{ADJ_NOTE}'].font = font_small

# ══════════════════════════════════════════════
# CHARTS — placed in rows 7-37 area
# ══════════════════════════════════════════════
# Chart 1: Total Revenue (Line) — top-left
chart1 = LineChart()
chart1.title = "Total Revenue by Scenario"
chart1.style = 2
chart1.y_axis.title = "CHF"
chart1.y_axis.numFmt = '#,##0'
chart1.x_axis.title = None
chart1.height = 10
chart1.width = 14.5
chart1.legend.position = 'b'
chart1.y_axis.delete = False
chart1.x_axis.delete = False
chart1.y_axis.majorGridlines = None

cats = Reference(ws, min_col=2, max_col=11, min_row=B_YR)

s1 = Reference(ws, min_col=2, max_col=11, min_row=base['rev'])
chart1.add_data(s1, from_rows=True)
chart1.series[0].title = openpyxl.chart.series.SeriesLabel(v="Base")
chart1.series[0].graphicalProperties.line.solidFill = NAVY
chart1.series[0].graphicalProperties.line.width = 28000

s2 = Reference(ws, min_col=2, max_col=11, min_row=down['rev'])
chart1.add_data(s2, from_rows=True)
chart1.series[1].title = openpyxl.chart.series.SeriesLabel(v="Downside")
chart1.series[1].graphicalProperties.line.solidFill = RED_ACCENT
chart1.series[1].graphicalProperties.line.dashStyle = "dash"
chart1.series[1].graphicalProperties.line.width = 20000

s3 = Reference(ws, min_col=2, max_col=11, min_row=up['rev'])
chart1.add_data(s3, from_rows=True)
chart1.series[2].title = openpyxl.chart.series.SeriesLabel(v="Upside")
chart1.series[2].graphicalProperties.line.solidFill = GREEN_ACCENT
chart1.series[2].graphicalProperties.line.dashStyle = "dash"
chart1.series[2].graphicalProperties.line.width = 20000

chart1.set_categories(cats)

chart1.series[0].dLbls = DataLabelList()
chart1.series[0].dLbls.showVal = True
chart1.series[0].dLbls.numFmt = '#,##0'
chart1.series[0].dLbls.showCatName = False
chart1.series[0].dLbls.showSerName = False

chart1.series[0].smooth = True
chart1.series[1].smooth = True
chart1.series[2].smooth = True

# Chart 2: EBIT (Bar chart) — top-right
chart2 = BarChart()
chart2.title = "EBIT by Scenario"
chart2.style = 2
chart2.y_axis.title = "CHF"
chart2.y_axis.numFmt = '#,##0'
chart2.height = 10
chart2.width = 14.5
chart2.legend.position = 'b'
chart2.y_axis.crossesAt = 0

cats2 = Reference(ws, min_col=2, max_col=11, min_row=B_YR)

s1 = Reference(ws, min_col=2, max_col=11, min_row=base['ebit'])
chart2.add_data(s1, from_rows=True)
chart2.series[0].title = openpyxl.chart.series.SeriesLabel(v="Base")
chart2.series[0].graphicalProperties.solidFill = NAVY

s2 = Reference(ws, min_col=2, max_col=11, min_row=down['ebit'])
chart2.add_data(s2, from_rows=True)
chart2.series[1].title = openpyxl.chart.series.SeriesLabel(v="Downside")
chart2.series[1].graphicalProperties.solidFill = RED_ACCENT

s3 = Reference(ws, min_col=2, max_col=11, min_row=up['ebit'])
chart2.add_data(s3, from_rows=True)
chart2.series[2].title = openpyxl.chart.series.SeriesLabel(v="Upside")
chart2.series[2].graphicalProperties.solidFill = GREEN_ACCENT

chart2.set_categories(cats2)

chart2.series[0].dLbls = DataLabelList()
chart2.series[0].dLbls.showVal = True
chart2.series[0].dLbls.numFmt = '#,##0'
chart2.series[0].dLbls.showCatName = False
chart2.series[0].dLbls.showSerName = False

# Chart 3: Cumulative Cash (Line) — bottom, full width
chart3 = LineChart()
chart3.title = "Cumulative Cash Flow by Scenario"
chart3.style = 2
chart3.y_axis.title = "CHF"
chart3.y_axis.numFmt = '#,##0'
chart3.height = 10
chart3.width = 29
chart3.legend.position = 'b'
chart3.y_axis.crossesAt = 0

cats3 = Reference(ws, min_col=2, max_col=11, min_row=B_YR)

s1 = Reference(ws, min_col=2, max_col=11, min_row=base['cum'])
chart3.add_data(s1, from_rows=True)
chart3.series[0].title = openpyxl.chart.series.SeriesLabel(v="Base")
chart3.series[0].graphicalProperties.line.solidFill = NAVY
chart3.series[0].graphicalProperties.line.width = 28000

s2 = Reference(ws, min_col=2, max_col=11, min_row=down['cum'])
chart3.add_data(s2, from_rows=True)
chart3.series[1].title = openpyxl.chart.series.SeriesLabel(v="Downside")
chart3.series[1].graphicalProperties.line.solidFill = RED_ACCENT
chart3.series[1].graphicalProperties.line.dashStyle = "dash"
chart3.series[1].graphicalProperties.line.width = 20000

s3 = Reference(ws, min_col=2, max_col=11, min_row=up['cum'])
chart3.add_data(s3, from_rows=True)
chart3.series[2].title = openpyxl.chart.series.SeriesLabel(v="Upside")
chart3.series[2].graphicalProperties.line.solidFill = GREEN_ACCENT
chart3.series[2].graphicalProperties.line.dashStyle = "dash"
chart3.series[2].graphicalProperties.line.width = 20000

chart3.set_categories(cats3)

chart3.series[0].dLbls = DataLabelList()
chart3.series[0].dLbls.showVal = True
chart3.series[0].dLbls.numFmt = '#,##0'
chart3.series[0].dLbls.showCatName = False
chart3.series[0].dLbls.showSerName = False

chart3.series[0].smooth = True
chart3.series[1].smooth = True
chart3.series[2].smooth = True

# Place charts: 2 side by side at top, 1 full-width below
ws.add_chart(chart1, "A7")
ws.add_chart(chart2, "F7")
ws.add_chart(chart3, "A22")

# ══════════════════════════════════════════════
# P&L SHEET: Rename "NET RESULT" to "PROFIT AFTER TAX"
# ══════════════════════════════════════════════
ws_pl = wb["P&L"]
ws_pl['A29'].value = "PROFIT AFTER TAX"

# ══════════════════════════════════════════════
# SHEET PROTECTION
# ══════════════════════════════════════════════
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

# Explicitly unlock Dashboard adjustment cells
for r in [ADJ_TICK, ADJ_PART, ADJ_COST]:
    for c in [2, 3, 4]:
        ws.cell(row=r, column=c).protection = Protection(locked=False)

# ══════════════════════════════════════════════
# PRINT SETTINGS
# ══════════════════════════════════════════════
ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.page_setup.orientation = 'landscape'

# ══════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════
wb.save(DST)
print("Done! Saved to", DST)
