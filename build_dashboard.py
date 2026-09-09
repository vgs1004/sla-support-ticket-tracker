"""
Adds real Excel PivotTables, PivotCharts and slicers to the Dashboard sheet
via COM automation (requires Excel installed on this machine).

Layout:
  - KPI cards: rows 4-6 (built in generate_data.py)
  - 2x2 PivotChart grid: starting ~row 10, fixed pixel positions (no overlap)
  - Slicers: to the right of the chart grid
  - Supporting PivotTables (the raw data behind each chart): parked well
    below, starting row 46, spread across separate column bands so none
    of them can collide regardless of how many rows/columns they grow to.
"""
import sys
import win32com.client as win32

sys.path.insert(0, r"C:\Users\varsh\Projects\Portfolio\_shared")
import palette as pal

PATH = r"C:\Users\varsh\Projects\Portfolio\02-sla-support-ticket-tracker-excel\SLA_Support_Ticket_Tracker.xlsx"

xlDatabase = 1
xlRowField = 1
xlColumnField = 2
xlDataField = 4
xlCount = -4112
xlColumnClustered = 51
xlPie = 5
xlLine = 4

xl = win32.gencache.EnsureDispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

wb = None
success = False
try:
    wb = xl.Workbooks.Open(PATH)
    xl.CalculateFullRebuild()
    ws_tickets = wb.Worksheets("Tickets")
    ws_dash = wb.Worksheets("Dashboard")

    used_range = ws_tickets.UsedRange
    last_row = used_range.Rows.Count  # includes header
    data_ref = "Tickets!$A$1:$L$" + str(last_row)  # A..L incl. MonthLogged helper col

    # Create a real Excel Table (ListObject) via COM -- this is what openpyxl's
    # Table object could not safely produce for automation (see generate_data.py note).
    lo = ws_tickets.ListObjects.Add(1, ws_tickets.Range(data_ref), None, 1)  # xlSrcRange=1, xlYes=1
    lo.Name = "TicketsTable"
    lo.TableStyle = "TableStyleMedium2"

    src = "Tickets!TicketsTable"
    cache1 = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=src)

    def make_pivot(cache, dest_cell, name, row_field=None, col_field=None, page_field=None):
        pt = cache.CreatePivotTable(TableDestination=ws_dash.Range(dest_cell), TableName=name)
        if row_field:
            pt.PivotFields(row_field).Orientation = xlRowField
        if col_field:
            pt.PivotFields(col_field).Orientation = xlColumnField
        if page_field:
            pt.PivotFields(page_field).Orientation = 3  # xlPageField
        df = pt.PivotFields("TicketID")
        df.Orientation = xlDataField
        df.Function = xlCount
        df.Caption = "Count of Tickets"
        return pt

    # Supporting PivotTables, each in its own column band so none can ever collide.
    PIVOT_ROW = 46
    pt1 = make_pivot(cache1, f"B{PIVOT_ROW}", "PT_TierPriority", row_field="Tier", col_field="Priority")
    pt2 = make_pivot(cache1, f"L{PIVOT_ROW}", "PT_SLAStatus", row_field="SLA_Status")
    try:
        pt2.PivotFields("SLA_Status").PivotItems("Open").Visible = False
    except Exception:
        pass
    pt3 = make_pivot(cache1, f"V{PIVOT_ROW}", "PT_BreachByPriority", row_field="Priority", page_field="SLA_Status")
    try:
        pt3.PivotFields("SLA_Status").CurrentPage = "Breached"
    except Exception:
        pass

    # Open vs Closed over time, by month -- uses the MonthLogged helper column
    # (more reliable across Excel/COM versions than PivotField date grouping).
    pt4 = cache1.CreatePivotTable(TableDestination=ws_dash.Range(f"AF{PIVOT_ROW}"), TableName="PT_OpenClosedTime")
    pt4.PivotFields("MonthLogged").Orientation = xlRowField
    pt4.PivotFields("Status").Orientation = xlColumnField
    df4 = pt4.PivotFields("TicketID")
    df4.Orientation = xlDataField
    df4.Function = xlCount
    df4.Caption = "Count of Tickets"

    dash_label = ws_dash.Range(f"B{PIVOT_ROW - 2}")
    dash_label.Value = "Supporting PivotTables (source data behind the 4 charts above)"
    dash_label.Font.Italic = True
    dash_label.Font.Size = 9
    dash_label.Font.Color = pal.rgb(pal.LABEL_GRAY)

    wb.Save()

    # ---- Charts: fixed 2x2 pixel grid, so nothing can overlap regardless
    # ---- of how the pivot tables above render.
    grid_top = ws_dash.Range("B10").Top
    grid_left = ws_dash.Range("B10").Left
    chart_w, chart_h = 430, 250
    gap_x, gap_y = 20, 20

    def style_chart(chart, title):
        chart.HasTitle = True
        chart.ChartTitle.Text = title
        chart.ChartTitle.Font.Color = pal.rgb(pal.INK)
        chart.ChartTitle.Font.Bold = True
        try:
            chart.ChartArea.Format.Line.Visible = False
        except Exception:
            pass

    def color_series(chart, colors):
        for i in range(1, chart.SeriesCollection().Count + 1):
            c = pal.rgb(colors[(i - 1) % len(colors)])
            s = chart.SeriesCollection(i)
            try:
                s.Format.Fill.ForeColor.RGB = c
            except Exception:
                pass
            try:
                s.Format.Line.ForeColor.RGB = c  # line charts: the stroke, not the fill
            except Exception:
                pass

    def color_pie_points(chart, colors, n):
        s = chart.SeriesCollection(1)
        for i in range(1, n + 1):
            try:
                s.Points(i).Format.Fill.ForeColor.RGB = pal.rgb(colors[(i - 1) % len(colors)])
            except Exception:
                pass

    def add_chart(pt, left, top, width, height, chart_type, title):
        co = ws_dash.ChartObjects().Add(Left=left, Top=top, Width=width, Height=height)
        co.Chart.SetSourceData(Source=pt.TableRange2)
        co.Chart.ChartType = chart_type
        style_chart(co.Chart, title)
        return co

    c1 = add_chart(pt1, grid_left, grid_top, chart_w, chart_h,
                    xlColumnClustered, "Tickets by Tier and Priority")
    color_series(c1.Chart, pal.CHART_SERIES)

    c2 = add_chart(pt2, grid_left + chart_w + gap_x, grid_top, chart_w, chart_h,
                    xlPie, "SLA Compliance: Met vs Breached")
    # Pivot hides "Open", leaving Breached then Met (alphabetical) as the two points
    color_pie_points(c2.Chart, [pal.ACCENT_RUST, pal.ACCENT_SAGE], 2)

    c3 = add_chart(pt3, grid_left, grid_top + chart_h + gap_y, chart_w, chart_h,
                    xlColumnClustered, "Breaches by Priority")
    color_series(c3.Chart, [pal.ACCENT_RUST])

    c4 = add_chart(pt4, grid_left + chart_w + gap_x, grid_top + chart_h + gap_y, chart_w, chart_h,
                    xlLine, "Open vs Closed Over Time")
    color_series(c4.Chart, pal.CHART_SERIES)

    wb.Save()

    # ---- Slicers for Priority and Tier, connected to all pivots, placed
    # ---- to the right of the chart grid.
    slicer_left = grid_left + 2 * chart_w + 2 * gap_x + 20
    slicer_top = grid_top

    sc_priority = wb.SlicerCaches.Add2(pt1, "Priority")
    sl1 = sc_priority.Slicers.Add(SlicerDestination=ws_dash, Name="PrioritySlicer", Caption="Priority",
                                   Top=slicer_top, Left=slicer_left, Width=150, Height=160)
    try:
        sl1.Style = "SlicerStyleLight3"  # warm/orange built-in style, not the default blue
    except Exception:
        pass
    for pt in (pt2, pt3, pt4):
        try:
            sc_priority.PivotTables.AddPivotTable(pt)
        except Exception as e:
            print("priority link skip:", e)

    sc_tier = wb.SlicerCaches.Add2(pt1, "Tier")
    sl2 = sc_tier.Slicers.Add(SlicerDestination=ws_dash, Name="TierSlicer", Caption="Tier",
                               Top=slicer_top + 170, Left=slicer_left, Width=150, Height=160)
    try:
        sl2.Style = "SlicerStyleLight3"
    except Exception:
        pass
    for pt in (pt2, pt4):
        try:
            sc_tier.PivotTables.AddPivotTable(pt)
        except Exception as e:
            print("tier link skip:", e)

    ws_dash.Activate()
    xl.ActiveWindow.View = 1  # xlNormalView
    ws_dash.Range("A1").Select()
    xl.ActiveWindow.Zoom = 85

    wb.Save()
    success = True
    print("Dashboard built.")
finally:
    if wb is not None:
        try:
            wb.Close(SaveChanges=success)
        except Exception:
            pass
    xl.Quit()
