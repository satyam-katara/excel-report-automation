import pandas as pd
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import ColorScaleRule


def export_report(
    df,
    kpis,
    category_report,
    subcategory_report,
    region_report,
    state_report,
    monthly_report,
    yearly_report,
    output_path,
    selected_month,
):
    """
    Create an eight-sheet Superstore MIS workbook with a modern corporate
    Executive Dashboard and a left-side filter/navigation panel.

    Note:
        OpenPyXL can create the dashboard layout and Excel charts, but it
        cannot create native interactive PivotTable slicers. The sidebar
        therefore provides filter guidance and report context; use the
        AutoFilter dropdowns on the underlying analysis sheets to filter
        their tables. Existing calculations and automation interfaces remain
        unchanged.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    selected_month = pd.Period(selected_month, freq="M")

    monthly_data = df[df["Year-Month"] == selected_month].copy()
    if monthly_data.empty:
        raise ValueError(f"No transactions found for {selected_month}")

    # Keep the same eight worksheets and the same function interface.
    sheets = {
        "Executive Dashboard": pd.DataFrame(
            list(kpis.items()), columns=["Metric", "Value"]
        ),
        "Category Analysis": category_report,
        "Subcategory Analysis": subcategory_report,
        "Regional Analysis": region_report,
        "State Analysis": state_report,
        "Monthly Summary": monthly_report,
        "Yearly Summary": yearly_report,
        "Monthly Data": monthly_data,
    }

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, data in sheets.items():
            export_df = data.copy()
            for col in export_df.columns:
                if isinstance(export_df[col].dtype, pd.PeriodDtype):
                    export_df[col] = export_df[col].astype(str)
            export_df.to_excel(writer, sheet_name=sheet_name, index=False)

    wb = load_workbook(output_path)

    # Modern corporate palette.
    NAVY = "142D4E"
    NAVY_DARK = "10243D"
    BLUE = "3977B8"
    TEAL = "159A9C"
    GREEN = "279B72"
    AMBER = "E5A13A"
    PURPLE = "8067B7"
    RED = "D76565"
    BG = "F3F6FA"
    WHITE = "FFFFFF"
    TEXT = "24364B"
    MUTED = "718096"
    BORDER = "DFE6EF"
    PALE_BLUE = "EAF2FB"
    PALE_TEAL = "E8F5F4"
    PALE_GREEN = "EAF5EF"
    PALE_AMBER = "FBF3E6"
    PALE_PURPLE = "F1EDFA"
    PALE_RED = "F9ECEC"

    thin_border = Border(
        left=Side(style="thin", color=BORDER),
        right=Side(style="thin", color=BORDER),
        top=Side(style="thin", color=BORDER),
        bottom=Side(style="thin", color=BORDER),
    )
    header_fill = PatternFill("solid", fgColor=NAVY)
    header_font = Font(name="Aptos", bold=True, color=WHITE, size=10)

    # Format all analytical/data sheets without changing their data.
    for report_ws in wb.worksheets:
        report_ws.sheet_view.showGridLines = False
        report_ws.freeze_panes = "A2"

        if report_ws.title != "Executive Dashboard":
            report_ws.auto_filter.ref = report_ws.dimensions

        for cell in report_ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.border = thin_border

        report_ws.row_dimensions[1].height = 26

        for column_cells in report_ws.columns:
            col_letter = column_cells[0].column_letter
            max_length = max(
                (
                    len(str(cell.value))
                    for cell in column_cells
                    if cell.value is not None
                ),
                default=10,
            )
            report_ws.column_dimensions[col_letter].width = min(
                max(max_length + 2, 12), 38
            )

            for cell in column_cells:
                cell.border = thin_border
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"

    # Dashboard canvas: A:D sidebar, E:X main content.
    ws = wb["Executive Dashboard"]
    ws.delete_rows(1, ws.max_row)
    ws._charts = []
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 95
    ws.sheet_view.zoomScaleNormal = 95
    ws.sheet_view.showRowColHeaders = False
    ws.freeze_panes = "E5"

    for col in ["A", "B"]:
        ws.column_dimensions[col].width = 4.8
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 13
    for col_idx in range(5, 25):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 12.2

    # Canvas background.
    for row in range(1, 39):
        ws.row_dimensions[row].height = 20
        for col_idx in range(1, 25):
            ws.cell(row=row, column=col_idx).fill = PatternFill(
                "solid", fgColor=BG
            )

    # Sidebar brand block.
    ws.merge_cells("A1:D3")
    ws["A1"] = "SUPERSTORE\nMIS"
    ws["A1"].font = Font(name="Aptos Display", size=19, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY_DARK)
    ws["A1"].alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    for row in ws["A1:D3"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=NAVY_DARK)

    # Sidebar section helper.
    def sidebar_section(row, title):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        cell = ws.cell(row=row, column=1)
        cell.value = title
        cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        for cells in ws.iter_rows(
            min_row=row, max_row=row, min_col=1, max_col=4
        ):
            for c in cells:
                c.fill = PatternFill("solid", fgColor=BLUE)
        ws.row_dimensions[row].height = 23

    def sidebar_value(row, label, value):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        left = ws.cell(row=row, column=1)
        right = ws.cell(row=row, column=3)
        left.value = label
        right.value = value
        left.font = Font(name="Aptos", size=9, color=MUTED)
        right.font = Font(name="Aptos", size=9, bold=True, color=TEXT)
        left.alignment = Alignment(vertical="center", wrap_text=True)
        right.alignment = Alignment(
            horizontal="right", vertical="center", wrap_text=True
        )
        for cells in ws.iter_rows(
            min_row=row, max_row=row, min_col=1, max_col=4
        ):
            for c in cells:
                c.fill = PatternFill("solid", fgColor=WHITE)
                c.border = thin_border
        ws.row_dimensions[row].height = 25

    sidebar_section(5, "REPORT CONTEXT")
    sidebar_value(6, "Reporting month", selected_month.strftime("%b %Y"))
    sidebar_value(7, "Dataset period", "2014 – 2017")
    sidebar_value(8, "Source records", f"{len(df):,}")
    sidebar_value(9, "Report rows", f"{len(monthly_data):,}")

    sidebar_section(11, "FILTER & EXPLORE")
    sidebar_instructions = [
        ("01", "Category Analysis"),
        ("02", "Subcategory Analysis"),
        ("03", "Regional Analysis"),
        ("04", "State Analysis"),
        ("05", "Monthly Summary"),
        ("06", "Yearly Summary"),
        ("07", "Monthly Data"),
    ]
    for idx, (number, label) in enumerate(sidebar_instructions, start=12):
        ws.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=1)
        ws.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=4)
        n = ws.cell(row=idx, column=1)
        n.value = number
        n.font = Font(name="Aptos", size=9, bold=True, color=BLUE)
        n.alignment = Alignment(horizontal="center", vertical="center")
        t = ws.cell(row=idx, column=2)
        t.value = label
        t.font = Font(name="Aptos", size=9, color=TEXT)
        t.alignment = Alignment(vertical="center")
        for cells in ws.iter_rows(
            min_row=idx, max_row=idx, min_col=1, max_col=4
        ):
            for c in cells:
                c.fill = PatternFill("solid", fgColor=WHITE)
                c.border = thin_border
        ws.row_dimensions[idx].height = 23

    sidebar_section(20, "HOW TO FILTER")
    ws.merge_cells("A21:D24")
    ws["A21"] = (
        "Open an analysis sheet and use the dropdown arrows in its header row "
        "to filter by Region, Category, Segment, State or Month. "
        "The dashboard charts are based on the full summary tables."
    )
    ws["A21"].font = Font(name="Aptos", size=9, color=TEXT)
    ws["A21"].alignment = Alignment(
        horizontal="left", vertical="top", wrap_text=True
    )
    for row in ws["A21:D24"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=WHITE)
            cell.border = thin_border

    sidebar_section(26, "ADDITIONAL KPIs")
    extra_kpis = [
        ("Customers", kpis.get("Total Customers", "N/A"), "#,##0"),
        ("Avg. order value", kpis.get("Average Order Value", "N/A"), "$#,##0.00"),
    ]
    for row, (label, value, fmt) in enumerate(extra_kpis, start=27):
        sidebar_value(row, label, value)
        if isinstance(value, (int, float)):
            ws.cell(row=row, column=3).number_format = fmt

    ws.merge_cells("A30:D32")
    ws["A30"] = "AUTOMATED MIS REPORT\nPython  |  Pandas  |  Excel  |  VBA"
    ws["A30"].font = Font(name="Aptos", size=10, bold=True, color=WHITE)
    ws["A30"].fill = PatternFill("solid", fgColor=NAVY_DARK)
    ws["A30"].alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    for row in ws["A30:D32"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=NAVY_DARK)

    # Main dashboard title.
    ws.merge_cells("E1:X2")
    ws["E1"] = "EXECUTIVE SALES DASHBOARD"
    ws["E1"].font = Font(name="Aptos Display", size=25, bold=True, color=WHITE)
    ws["E1"].fill = PatternFill("solid", fgColor=NAVY_DARK)
    ws["E1"].alignment = Alignment(horizontal="left", vertical="center")
    for row in ws["E1:X2"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=NAVY_DARK)

    ws.merge_cells("E3:X3")
    ws["E3"] = (
        f"Reporting period: {selected_month.strftime('%B %Y')}   |   "
        "Business overview: complete dataset (2014–2017)"
    )
    ws["E3"].font = Font(name="Aptos", size=10, color=MUTED, italic=True)
    ws["E3"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 9

    # Six KPI cards across the main dashboard area.
    kpi_cards = [
        ("E5:G5", "E6:G7", "TOTAL SALES", kpis["Total Sales"],
         '$#,##0', BLUE, PALE_BLUE),
        ("H5:J5", "H6:J7", "TOTAL PROFIT", kpis["Total Profit"],
         '$#,##0', GREEN, PALE_GREEN),
        ("K5:M5", "K6:M7", "TOTAL ORDERS", kpis["Total Orders"],
         '#,##0', AMBER, PALE_AMBER),
        ("N5:P5", "N6:P7", "PROFIT MARGIN", kpis["Profit Margin (%)"],
         '0.00"%"', PURPLE, PALE_PURPLE),
        ("Q5:S5", "Q6:S7", "CUSTOMERS", kpis["Total Customers"],
         '#,##0', TEAL, PALE_TEAL),
        ("T5:X5", "T6:X7", "AVG ORDER VALUE", kpis["Average Order Value"],
         '$#,##0.00', RED, PALE_RED),
    ]

    for label_range, value_range, label, value, number_format, accent, pale in kpi_cards:
        ws.merge_cells(label_range)
        ws.merge_cells(value_range)
        label_cell = ws[label_range.split(":")[0]]
        value_cell = ws[value_range.split(":")[0]]

        label_cell.value = label
        label_cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
        label_cell.fill = PatternFill("solid", fgColor=accent)
        label_cell.alignment = Alignment(horizontal="center", vertical="center")

        value_cell.value = value
        value_cell.number_format = number_format
        value_cell.font = Font(
            name="Aptos Display", size=16, bold=True, color=NAVY_DARK
        )
        value_cell.fill = PatternFill("solid", fgColor=WHITE)
        value_cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in ws[label_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=accent)
        for row in ws[value_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=WHITE)
                cell.border = thin_border

    ws.row_dimensions[5].height = 23
    ws.row_dimensions[6].height = 28
    ws.row_dimensions[7].height = 25
    ws.row_dimensions[8].height = 10

    def find_column(sheet, candidates):
        headers = {
            str(cell.value).strip().lower(): cell.column
            for cell in sheet[1]
            if cell.value is not None
        }
        for candidate in candidates:
            if candidate.lower() in headers:
                return headers[candidate.lower()]
        raise ValueError(
            f"Could not find any of {candidates} in sheet '{sheet.title}'. "
            f"Available columns: {list(headers)}"
        )

    def panel_heading(title, cell_range):
        ws.merge_cells(cell_range)
        cell = ws[cell_range.split(":")[0]]
        cell.value = title
        cell.font = Font(name="Aptos", size=12, bold=True, color=NAVY)
        cell.fill = PatternFill("solid", fgColor=WHITE)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        for row in ws[cell_range]:
            for c in row:
                c.fill = PatternFill("solid", fgColor=WHITE)
                c.border = Border(bottom=Side(style="medium", color=BLUE))

    # Chart panels: two charts per row.
    panel_heading("MONTHLY SALES TREND", "E9:N9")
    panel_heading("MONTHLY PROFIT TREND", "O9:X9")
    panel_heading("SALES BY CATEGORY", "E20:N20")
    panel_heading("SALES BY REGION", "O20:X20")
    ws.row_dimensions[9].height = 26
    ws.row_dimensions[20].height = 26

    monthly_ws = wb["Monthly Summary"]
    month_col = find_column(monthly_ws, ["Year-Month", "Month", "Order Month"])
    sales_col = find_column(monthly_ws, ["Total_Sales", "Total Sales", "Sales"])
    profit_col = find_column(monthly_ws, ["Total_Profit", "Total Profit", "Profit"])

    month_labels = Reference(
        monthly_ws, min_col=month_col, min_row=2, max_row=monthly_ws.max_row
    )

    sales_chart = LineChart()
    sales_chart.title = None
    sales_chart.style = 13
    sales_chart.height = 6.2
    sales_chart.width = 19.8
    sales_chart.legend = None
    sales_chart.y_axis.title = "Sales ($)"
    sales_chart.x_axis.title = "Month"
    sales_chart.y_axis.numFmt = '$#,##0,"K"'
    sales_chart.x_axis.tickLblPos = "low"
    sales_chart.x_axis.tickLblSkip = 3
    sales_chart.x_axis.tickMarkSkip = 3
    sales_chart.display_blanks = "gap"
    sales_data = Reference(
        monthly_ws, min_col=sales_col, min_row=1, max_row=monthly_ws.max_row
    )
    sales_chart.add_data(sales_data, titles_from_data=True)
    sales_chart.set_categories(month_labels)
    sales_chart.series[0].graphicalProperties.line.solidFill = BLUE
    sales_chart.series[0].graphicalProperties.line.width = 28000
    sales_chart.series[0].marker.symbol = "circle"
    sales_chart.series[0].marker.size = 4
    ws.add_chart(sales_chart, "E10")

    profit_chart = BarChart()
    profit_chart.type = "col"
    profit_chart.grouping = "clustered"
    profit_chart.overlap = 0
    profit_chart.title = None
    profit_chart.style = 10
    profit_chart.height = 6.2
    profit_chart.width = 19.8
    profit_chart.legend = None
    profit_chart.y_axis.title = "Profit ($)"
    profit_chart.x_axis.title = "Month"
    profit_chart.y_axis.numFmt = '$#,##0,"K"'
    profit_chart.x_axis.tickLblSkip = 3
    profit_chart.x_axis.tickMarkSkip = 3
    profit_data = Reference(
        monthly_ws, min_col=profit_col, min_row=1, max_row=monthly_ws.max_row
    )
    profit_chart.add_data(profit_data, titles_from_data=True)
    profit_chart.set_categories(month_labels)
    profit_chart.series[0].graphicalProperties.solidFill = GREEN
    profit_chart.series[0].graphicalProperties.line.solidFill = GREEN
    ws.add_chart(profit_chart, "O10")

    category_ws = wb["Category Analysis"]
    category_name_col = find_column(category_ws, ["Category"])
    category_sales_col = find_column(
        category_ws, ["Total_Sales", "Total Sales", "Sales"]
    )

    category_chart = DoughnutChart()
    category_chart.title = None
    category_chart.style = 10
    category_chart.holeSize = 62
    category_chart.height = 5.8
    category_chart.width = 19.8
    category_chart.legend.position = "r"
    category_chart.firstSliceAng = 270
    category_data = Reference(
        category_ws, min_col=category_sales_col, min_row=1,
        max_row=category_ws.max_row
    )
    category_labels = Reference(
        category_ws, min_col=category_name_col, min_row=2,
        max_row=category_ws.max_row
    )
    category_chart.add_data(category_data, titles_from_data=True)
    category_chart.set_categories(category_labels)
    category_chart.dataLabels = DataLabelList()
    category_chart.dataLabels.showVal = False
    category_chart.dataLabels.showPercent = False
    category_chart.dataLabels.showCatName = False
    category_chart.dataLabels.showSerName = False
    category_chart.dataLabels.showLegendKey = False

    for series in category_chart.series:
        for idx, color in enumerate([BLUE, AMBER, TEAL]):
            series.data_points.append(
                DataPoint(
                    idx=idx,
                    spPr=GraphicalProperties(solidFill=color)
                )
            )
    ws.add_chart(category_chart, "E21")

    region_ws = wb["Regional Analysis"]
    region_name_col = find_column(region_ws, ["Region"])
    region_sales_col = find_column(
        region_ws, ["Total_Sales", "Total Sales", "Sales"]
    )

    region_chart = BarChart()
    region_chart.type = "bar"
    region_chart.grouping = "clustered"
    region_chart.overlap = 100
    region_chart.title = None
    region_chart.style = 12
    region_chart.height = 5.8
    region_chart.width = 19.8
    region_chart.legend = None
    region_chart.x_axis.title = "Sales ($)"
    region_chart.y_axis.title = "Region"
    region_chart.x_axis.numFmt = '$#,##0,"K"'
    region_chart.y_axis.reverseOrder = True
    region_data = Reference(
        region_ws, min_col=region_sales_col, min_row=1,
        max_row=region_ws.max_row
    )
    region_labels = Reference(
        region_ws, min_col=region_name_col, min_row=2,
        max_row=region_ws.max_row
    )
    region_chart.add_data(region_data, titles_from_data=True)
    region_chart.set_categories(region_labels)
    region_chart.series[0].graphicalProperties.solidFill = BLUE
    region_chart.series[0].graphicalProperties.line.solidFill = BLUE
    region_chart.dataLabels = DataLabelList()
    region_chart.dataLabels.showVal = True
    region_chart.dataLabels.showCatName = False
    region_chart.dataLabels.showSerName = False
    region_chart.dataLabels.showLegendKey = False
    region_chart.dataLabels.position = "outEnd"
    region_chart.dataLabels.numFmt = '$#,##0'
    ws.add_chart(region_chart, "O21")

    # Profit/loss conditional formatting remains on analysis sheets.
    for sheet_name in [
        "Category Analysis",
        "Subcategory Analysis",
        "Regional Analysis",
        "State Analysis",
    ]:
        report_ws = wb[sheet_name]
        headers = {
            str(cell.value).strip(): cell.column
            for cell in report_ws[1]
            if cell.value is not None
        }
        profit_header = next(
            (
                name
                for name in ["Total_Profit", "Total Profit", "Profit"]
                if name in headers
            ),
            None,
        )
        if profit_header:
            col_letter = report_ws.cell(
                row=1, column=headers[profit_header]
            ).column_letter
            report_ws.conditional_formatting.add(
                f"{col_letter}2:{col_letter}{report_ws.max_row}",
                ColorScaleRule(
                    start_type="min",
                    start_color="F8696B",
                    mid_type="percentile",
                    mid_value=50,
                    mid_color="FFEB84",
                    end_type="max",
                    end_color="63BE7B",
                ),
            )

    # Print layout and workbook opening position.
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = "landscape"
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3
    ws.print_area = "A1:X36"
    ws.sheet_properties.outlinePr.summaryBelow = True

    wb.active = wb.sheetnames.index("Executive Dashboard")
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False

    wb.save(output_path)

    print("Excel MIS report generated successfully!")
    print("File:", output_path)
    print("Selected Month:", selected_month)
    print("Sheets:", len(wb.sheetnames))

    return output_path
