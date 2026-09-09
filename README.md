# SLA and Support Ticket Tracker

An Excel-based SLA tracker for a support desk: raw ticket log, automatic
breach calculation, and a PivotTable/PivotChart dashboard with slicers.
Built to mirror day-to-day L1/L2/L3 SLA tracking in a fully anonymized,
synthetic form.

## What it is

- 180 synthetic support tickets across P1 to P4 priorities, three support
  tiers (L1/L2/L3), six categories, and a full 2025 calendar year.
- No real company data. All ticket content, dates, and owner names are
  generated (see `generate_data.py`).

## Tools used

- Excel (formulas, data validation, conditional formatting, native
  PivotTables/PivotCharts, slicers)
- Python (`openpyxl`, `pywin32`) to generate the data and assemble the
  workbook programmatically. Scripts are included so the build is
  reproducible from scratch.

## What it shows

**Tickets sheet**
- Dropdown validation for Priority, Category, Tier, and Status.
- `ResolutionHours` and `SLA_Status` computed live from `DateLogged`,
  `ResolvedDateTime`, and each priority's SLA target (P1 = 4h, P2 = 8h,
  P3 = 24h, P4 = 72h).
- Conditional formatting (RAG): green for Met, red for Breached, amber for
  still Open.

**Dashboard sheet**
- KPI cards: Total Tickets, SLA Compliance %, Open Tickets, Total Breaches,
  Average Resolution Hours.
- Four PivotCharts: Tickets by Tier and Priority, SLA Compliance (Met vs
  Breached), Breaches by Priority, and Open vs Closed volume by month.
- Priority and Tier slicers that cross-filter all four charts at once.
- The native PivotTables backing each chart are included below the charts
  for transparency.

## Files

| File | Purpose |
|---|---|
| `SLA_Support_Ticket_Tracker.xlsx` | The finished workbook |
| `sample_tickets_data.csv` | Standalone copy of the raw ticket data |
| `generate_data.py` | Generates the synthetic tickets and builds the Tickets/KPI layout |
| `build_dashboard.py` | Adds the native PivotTables, PivotCharts, and slicers via Excel automation |
| `screenshots/` | Dashboard and data views (see below, since GitHub can't preview `.xlsx`) |

## Screenshots

- `01_dashboard_overview.png`: KPI cards and the four PivotCharts
- `02_dashboard_pivot_tables.png`: the supporting PivotTables behind the charts
- `03_tickets_sheet.png`: raw ticket log with RAG conditional formatting
