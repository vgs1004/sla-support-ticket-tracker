"""
Generate synthetic SLA / support ticket data and build the Tickets sheet
(raw data + formulas + validation + conditional formatting) with openpyxl.
The Dashboard sheet (pivots/charts/slicers) is added afterwards via COM
in build_dashboard.py, because openpyxl cannot create native PivotTables.
"""
import random
import sys
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

sys.path.insert(0, r"C:\Users\varsh\Projects\Portfolio\_shared")
import palette as pal

random.seed(42)

N = 180
START = datetime(2025, 1, 1)
END = datetime(2025, 12, 31)

PRIORITIES = ["P1", "P2", "P3", "P4"]
PRIORITY_WEIGHTS = [0.08, 0.22, 0.45, 0.25]
SLA_TARGET = {"P1": 4, "P2": 8, "P3": 24, "P4": 72}
CATEGORIES = ["Access", "Bug", "Data", "Change", "Hardware", "Network"]
TIERS = ["L1", "L2", "L3"]
TIER_WEIGHTS = [0.55, 0.32, 0.13]
STATUSES_OPEN = ["Open", "In Progress"]
OWNERS = [
    "A. Mehta", "R. Fernandes", "S. Iyer", "K. Padilla", "J. Ncube",
    "L. Duarte", "P. Osei", "M. Choudhury", "T. Alvarez", "N. Kowalski",
]


def random_datetime(start, end):
    delta = end - start
    seconds = random.randint(0, int(delta.total_seconds()))
    return start + timedelta(seconds=seconds)


rows = []
for i in range(1, N + 1):
    ticket_id = f"TCK-{i:04d}"
    priority = random.choices(PRIORITIES, weights=PRIORITY_WEIGHTS)[0]
    category = random.choice(CATEGORIES)
    tier = random.choices(TIERS, weights=TIER_WEIGHTS)[0]
    owner = random.choice(OWNERS)
    date_logged = random_datetime(START, END)
    sla_target = SLA_TARGET[priority]

    # ~78% of tickets are resolved; rest stay open/in progress (mostly recent ones)
    is_recent = date_logged > END - timedelta(days=10)
    resolved = random.random() < 0.78 and not (is_recent and random.random() < 0.5)

    resolved_dt = ""
    if resolved:
        # Simulate resolution time: usually near/under target, sometimes breached
        breach_chance = {"P1": 0.22, "P2": 0.20, "P3": 0.18, "P4": 0.15}[priority]
        if random.random() < breach_chance:
            hours = sla_target * random.uniform(1.05, 2.8)
        else:
            hours = sla_target * random.uniform(0.05, 0.98)
        resolved_dt = date_logged + timedelta(hours=hours)
        if resolved_dt > END:
            resolved_dt = END - timedelta(hours=random.uniform(1, 24))
        status = "Closed" if random.random() < 0.65 else "Resolved"
    else:
        status = random.choice(STATUSES_OPEN)

    rows.append(
        [
            ticket_id,
            date_logged,
            priority,
            category,
            tier,
            status,
            sla_target,
            resolved_dt,
            owner,
        ]
    )

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Tickets"

headers = [
    "TicketID", "DateLogged", "Priority", "Category", "Tier", "Status",
    "SLA_Target_Hours", "ResolvedDateTime", "ResolutionHours", "SLA_Status", "Owner",
    "MonthLogged",
]
ws.append(headers)

header_fill = PatternFill("solid", fgColor=pal.INK)
header_font = Font(color="FFFFFF", bold=True)
for col in range(1, len(headers) + 1):
    c = ws.cell(row=1, column=col)
    c.fill = header_fill
    c.font = header_font
    c.alignment = Alignment(horizontal="center", vertical="center")

for r, row in enumerate(rows, start=2):
    (ticket_id, date_logged, priority, category, tier, status,
     sla_target, resolved_dt, owner) = row
    ws.cell(row=r, column=1, value=ticket_id)
    ws.cell(row=r, column=2, value=date_logged).number_format = "yyyy-mm-dd hh:mm"
    ws.cell(row=r, column=3, value=priority)
    ws.cell(row=r, column=4, value=category)
    ws.cell(row=r, column=5, value=tier)
    ws.cell(row=r, column=6, value=status)
    ws.cell(row=r, column=7, value=sla_target)
    if resolved_dt:
        ws.cell(row=r, column=8, value=resolved_dt).number_format = "yyyy-mm-dd hh:mm"
    # ResolutionHours formula (blank if not resolved yet)
    ws.cell(row=r, column=9,
            value=f'=IF(H{r}="","",(H{r}-B{r})*24)').number_format = "0.0"
    # SLA_Status formula
    ws.cell(row=r, column=10,
            value=f'=IF(H{r}="","Open",IF(I{r}<=G{r},"Met","Breached"))')
    ws.cell(row=r, column=11, value=owner)
    # Helper column: first-of-month date, used as a clean pivot row field
    # for the "Open vs Closed Over Time" chart (more reliable than COM
    # pivot-field date grouping).
    ws.cell(row=r, column=12, value=f"=DATE(YEAR(B{r}),MONTH(B{r}),1)").number_format = "mmm-yyyy"

# Column widths
widths = [11, 18, 9, 11, 7, 12, 17, 18, 15, 12, 14, 12]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

last_row = N + 1

# Data validation dropdowns
dv_priority = DataValidation(type="list", formula1='"P1,P2,P3,P4"', allow_blank=False)
dv_category = DataValidation(type="list", formula1=f'"{",".join(CATEGORIES)}"', allow_blank=False)
dv_tier = DataValidation(type="list", formula1='"L1,L2,L3"', allow_blank=False)
dv_status = DataValidation(type="list", formula1='"Open,In Progress,Resolved,Closed"', allow_blank=False)
for dv in (dv_priority, dv_category, dv_tier, dv_status):
    ws.add_data_validation(dv)
dv_priority.add(f"C2:C{last_row}")
dv_category.add(f"D2:D{last_row}")
dv_tier.add(f"E2:E{last_row}")
dv_status.add(f"F2:F{last_row}")

# Conditional formatting RAG on SLA_Status (col J)
green_fill = PatternFill("solid", fgColor=pal.MET_FILL)
green_font = Font(color=pal.MET_FONT)
red_fill = PatternFill("solid", fgColor=pal.BREACH_FILL)
red_font = Font(color=pal.BREACH_FONT)
amber_fill = PatternFill("solid", fgColor=pal.OPEN_FILL)
amber_font = Font(color=pal.OPEN_FONT)

rng = f"J2:J{last_row}"
ws.conditional_formatting.add(
    rng, CellIsRule(operator="equal", formula=['"Met"'], fill=green_fill, font=green_font)
)
ws.conditional_formatting.add(
    rng, CellIsRule(operator="equal", formula=['"Breached"'], fill=red_fill, font=red_font)
)
ws.conditional_formatting.add(
    rng, CellIsRule(operator="equal", formula=['"Open"'], fill=amber_fill, font=amber_font)
)

# Freeze header row, add autofilter
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:L{last_row}"

# NOTE: deliberately NOT using openpyxl's Table object here -- it produces
# workbooks that Excel's COM automation (Workbooks.Open with alerts
# suppressed) refuses to open, even though they open fine in the Excel UI
# with auto-repair. A real Excel ListObject/Table is added via COM instead
# in build_dashboard.py, and PivotTables source from it.

# Dashboard sheet: title + KPI cards (formulas). PivotTables/charts/slicers
# are added afterwards via COM in build_dashboard.py.
dash = wb.create_sheet("Dashboard")
dash.sheet_view.showGridLines = False

dash["B2"] = "SLA and Support Ticket Tracker: Dashboard"
dash["B2"].font = Font(size=16, bold=True, color=pal.INK)

kpi_labels = ["Total Tickets", "SLA Compliance %", "Open Tickets", "Total Breaches", "Avg Resolution Hours"]
kpi_formulas = [
    f"=COUNTA(Tickets!A2:A{last_row})",
    f'=COUNTIF(Tickets!J2:J{last_row},"Met")/(COUNTIF(Tickets!J2:J{last_row},"Met")+COUNTIF(Tickets!J2:J{last_row},"Breached"))',
    f'=COUNTIF(Tickets!J2:J{last_row},"Open")',
    f'=COUNTIF(Tickets!J2:J{last_row},"Breached")',
    f'=AVERAGEIF(Tickets!I2:I{last_row},"<>",Tickets!I2:I{last_row})',
]
kpi_fmt = [None, "0.0%", None, None, "0.0"]

card_fill = PatternFill("solid", fgColor=pal.CARD_FILL)
label_font = Font(size=10, color=pal.LABEL_GRAY)
value_font = Font(size=20, bold=True, color=pal.ACCENT_RUST)
thin = Side(style="thin", color=pal.CARD_BORDER)
box_border = Border(left=thin, right=thin, top=thin, bottom=thin)

# One card = one (wide) column, with a narrow spacer column after it.
# No merged cells -- merged numeric cells are unreliable for width-fit
# rendering, so each card gets a single column wide enough on its own.
card_col_width = 20
spacer_col_width = 8
start_col = 2  # column B

for i, (label, formula, fmt) in enumerate(zip(kpi_labels, kpi_formulas, kpi_fmt)):
    col = start_col + i * 2  # card col, then a spacer col
    col_l = get_column_letter(col)
    lbl_cell = dash[f"{col_l}4"]
    lbl_cell.value = label
    lbl_cell.font = label_font
    lbl_cell.alignment = Alignment(horizontal="left", vertical="center")
    val_cell = dash[f"{col_l}5"]
    val_cell.value = formula
    if fmt:
        val_cell.number_format = fmt
    val_cell.font = value_font
    val_cell.alignment = Alignment(horizontal="left", vertical="center")
    for r in (4, 5, 6):
        dash.cell(row=r, column=col).fill = card_fill
        dash.cell(row=r, column=col).border = box_border
    dash.row_dimensions[6].height = 6  # thin bottom pad row of the card

dash.column_dimensions["A"].width = 2
for i in range(5):
    col = start_col + i * 2
    dash.column_dimensions[get_column_letter(col)].width = card_col_width
    dash.column_dimensions[get_column_letter(col + 1)].width = spacer_col_width

dash["B8"] = "Charts above are PivotCharts with Priority + Tier slicers. Supporting PivotTables are below (row 40+)."
dash["B8"].font = Font(size=9, italic=True, color=pal.LABEL_GRAY)

out_path = r"C:\Users\varsh\Projects\Portfolio\02-sla-support-ticket-tracker-excel\SLA_Support_Ticket_Tracker.xlsx"
wb.save(out_path)
print("Saved", out_path, "rows:", N)
