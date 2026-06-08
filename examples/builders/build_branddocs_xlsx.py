# SPDX-License-Identifier: MIT
"""Deterministic builder for the COMPLEX synthetic XLSX example template.

Produces ``examples/templates/branddocs_template.xlsx``: a 100% synthetic
(BrandDocs-style, never proprietary) workbook that stresses the brand-xlsx engine
across as many Excel component types as openpyxl can author:

  * MULTIPLE sheets in a deliberate tab order (Cover, Inputs, Model, Summary,
    Dashboard, Scenarios, Data) - exercises multi-sheet structure + skeleton ordering.
  * NAMED regions of every geometry the extractor distinguishes:
      - a single-cell title anchor sitting under a MERGED header row,
      - a single-cell that sits in a FROZEN header band,
      - a multi-cell INPUTS block (sample-data candidate),
      - a multi-cell DATA region that is the BODY of a native table object,
      - a single named cell that is a cross-sheet formula target.
  * FORMULAS: in-sheet SUM/total rows, percent-of-total ratios, a cross-sheet
    reference (Summary pulls from Model and Inputs), and a SUBTOTAL - all
    authored verbatim so we can later assert they survive generation byte-exact.
  * NUMBER FORMATS: currency (accounting), percent, thousands, and ISO date.
  * A native TABLE object (``BrandDocsDataTbl``) with a banded table style.
  * A dashboard sheet with KPI cards, cross-sheet formulas, and a third native
    chart so the rendered workbook looks like a finished executive model.
  * A scenarios sheet with list validation and a native scenario table.
  * CONDITIONAL FORMATTING: a color scale, a cell-is rule, and a formula rule.
  * FROZEN PANES on the data sheets.
  * The shared text-only BrandDocs wordmark placed as a worksheet drawing on the
    cover - rendered in-process by ``_brandlib.branddocs_mark_png`` (no
    external/proprietary asset on disk), sized at 4:1 so the text is not
    distorted.
  * Named cell STYLES (``BrandDocsTitle``, ``BrandDocsHeader``, ``BrandDocsCurrency``,
    ``BrandDocsPercent``, ``BrandDocsInput``) registered on the workbook and applied.
  * Demo / sample data rows the generator is expected to clear.

The output is byte-reproducible: a fixed timestamp / fixed image bytes / sorted
defined names, so re-running the builder yields an identical file (CI-friendly).

Run:
    PYTHONPATH=scripts .venv/bin/python examples/builders/build_branddocs_xlsx.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
)
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    NamedStyle,
    PatternFill,
    Protection,
    Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from _brandlib import branddocs_mark_png, freeze_ooxml
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = Path(__file__).resolve().parents[1] / "templates" / "branddocs_template.xlsx"

# Synthetic BrandDocs brand palette (made-up; never proprietary).
BRAND_NAVY = "FF16213F"
BRAND_TEAL = "FF2B7CD3"
BRAND_AMBER = "FFE0742B"
BRAND_LIGHT = "FFEAF1FF"
WHITE = "FFFFFFFF"

# 6-digit (no alpha) variants for openpyxl chart series fills, derived from the
# brand palette constants above (no new brand literals introduced).
NAVY6 = BRAND_NAVY[-6:]  # "16213F"
TEAL6 = BRAND_TEAL[-6:]  # "2B7CD3"
AMBER6 = BRAND_AMBER[-6:]  # "E0742B"


# ---------------------------------------------------------------------------
# Named cell styles (registered once on the workbook).
# ---------------------------------------------------------------------------
def _register_named_styles(wb: Workbook) -> None:
    thin = Side(style="thin", color=BRAND_NAVY)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    title = NamedStyle(name="BrandDocsTitle")
    title.font = Font(name="Arial", size=20, bold=True, color=BRAND_NAVY)
    title.alignment = Alignment(horizontal="center", vertical="center")

    header = NamedStyle(name="BrandDocsHeader")
    header.font = Font(name="Arial", size=11, bold=True, color=WHITE)
    header.fill = PatternFill("solid", fgColor=BRAND_NAVY)
    header.alignment = Alignment(horizontal="center", vertical="center")
    header.border = border

    currency = NamedStyle(name="BrandDocsCurrency")
    currency.number_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    currency.font = Font(name="Calibri", size=11, color=BRAND_NAVY)
    currency.border = border

    percent = NamedStyle(name="BrandDocsPercent")
    percent.number_format = "0.0%"
    percent.font = Font(name="Calibri", size=11, color=BRAND_TEAL)
    percent.border = border

    inp = NamedStyle(name="BrandDocsInput")
    inp.fill = PatternFill("solid", fgColor=BRAND_LIGHT)
    inp.font = Font(name="Calibri", size=11, color=BRAND_NAVY)
    inp.border = border

    kpi = NamedStyle(name="BrandDocsKPI")
    kpi.fill = PatternFill("solid", fgColor=BRAND_NAVY)
    kpi.font = Font(name="Arial", size=15, bold=True, color=WHITE)
    kpi.alignment = Alignment(horizontal="center", vertical="center")
    kpi.border = border

    for style in (title, header, currency, percent, inp, kpi):
        wb.add_named_style(style)


# ---------------------------------------------------------------------------
# Sheet builders.
# ---------------------------------------------------------------------------
def _build_cover(wb: Workbook) -> None:
    ws = wb.create_sheet("Cover")
    ws.sheet_view.showGridLines = False
    # Merged title band A1:G1 with a single-cell named title anchor at A1.
    ws.merge_cells("A1:G1")
    # Title band: navy fill + white title text so the cover reads as branded.
    ws["A1"] = "FY2025 Revenue Performance Review"
    ws["A1"].style = "BrandDocsTitle"
    ws["A1"].font = Font(name="Arial", size=20, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=BRAND_NAVY)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:G2")
    ws["A2"] = "Quarterly revenue model and executive summary"
    ws["A2"].font = Font(name="Arial", size=12, italic=True, color=BRAND_TEAL)
    ws["A2"].alignment = Alignment(horizontal="center")
    ws["A4"] = "Prepared for"
    ws["B4"] = "BrandDocs Corp (synthetic demo)"
    ws["A5"] = "Reporting period"
    ws["B5"] = "FY2025 (Q1-Q4)"
    # An ISO-date cell with a date number format (pinned, deterministic).
    ws["A6"] = "Generated on"
    ws["B6"] = "2026-01-15"
    ws["B6"].number_format = "yyyy-mm-dd"
    # Three scorecard tiles make the cover read as an executive-ready workbook.
    scorecards = [
        ("D4", "FY net revenue", "=Summary!B4"),
        ("E4", "Net margin", "=Summary!B8"),
        ("F4", "Brand audit", "Deep QA"),
    ]
    for anchor, label, value in scorecards:
        col = ws[anchor].column
        ws.cell(row=4, column=col, value=label).style = "BrandDocsHeader"
        vc = ws.cell(row=5, column=col, value=value)
        vc.style = "BrandDocsKPI"
        if value.startswith("="):
            vc.number_format = "0.0%" if "B8" in value else "#,##0"
    ws["G4"] = "Scope"
    ws["G4"].style = "BrandDocsHeader"
    ws["G5"] = "3 formats"
    ws["G5"].style = "BrandDocsKPI"
    ws["G5"].alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    # A navy brand band across the printable width, below the scorecards.
    ws.merge_cells("A8:G8")
    for col in range(1, 8):
        ws.cell(row=8, column=col).fill = PatternFill("solid", fgColor=BRAND_NAVY)
    ws.row_dimensions[8].height = 28
    # Header drawing: the shared BrandDocs wordmark sits inside the printable
    # area at a 4:1 aspect ratio so the text is not distorted.
    logo = XLImage(_logo_path())
    logo.width, logo.height = 180, 45
    ws.add_image(logo, "A9")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 28
    for col in "CDEFG":
        ws.column_dimensions[col].width = 16
    for row in (4, 5):
        ws.row_dimensions[row].height = 28
    # Keep the whole cover on a single print page.
    ws.print_area = "A1:G13"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _build_inputs(wb: Workbook) -> None:
    ws = wb.create_sheet("Inputs")
    ws["A1"] = "Model Inputs"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color=BRAND_NAVY)
    # Header row 3.
    for col, label in enumerate(("Driver", "Value", "Unit"), start=1):
        c = ws.cell(row=3, column=col, value=label)
        c.style = "BrandDocsHeader"
    # Input block rows 4..7 (the named INPUTS region, demo values).
    # Reconciled so Implied gross (Units * Price) == Model Gross revenue FY
    # (29070 * 50 = 1,453,500), giving Net margin = 85.0% (sensible range).
    inputs = [
        ("Units sold", 29070, "ea"),
        ("Unit price", 50.0, "USD"),
        ("Discount rate", 0.12, "pct"),
        ("Tax rate", 0.22, "pct"),
    ]
    for i, (driver, val, unit) in enumerate(inputs):
        r = 4 + i
        ws.cell(row=r, column=1, value=driver).style = "BrandDocsInput"
        vc = ws.cell(row=r, column=2, value=val)
        vc.style = "BrandDocsInput"
        if unit == "USD":
            vc.style = "BrandDocsCurrency"
        elif unit == "pct":
            vc.style = "BrandDocsPercent"
        elif unit == "ea":
            # Whole-unit counts get a thousands separator so "Units sold" reads
            # "29,070" instead of "29070" (keeps the input fill/border style).
            vc.number_format = "#,##0"
        ws.cell(row=r, column=3, value=unit).style = "BrandDocsInput"
    ws.freeze_panes = "A4"
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    # Distributed-template convention: lock the sheet but leave ONLY the input
    # value cells (B4:B7) editable, so a recipient fills the blue inputs and cannot
    # disturb the model. (openpyxl honors locked/unlocked on save; enforcement is
    # Excel-side, so the generator can still fill the named region programmatically.)
    for r in range(4, 8):
        ws.cell(row=r, column=2).protection = Protection(locked=False)
    ws.protection.sheet = True
    # Print-ready header/footer with live fields (sheet name + page x of y).
    ws.oddHeader.center.text = "&A"
    ws.oddFooter.right.text = "Page &P of &N"


def _build_model(wb: Workbook) -> None:
    ws = wb.create_sheet("Model")
    ws["A1"] = "Revenue Model"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color=BRAND_NAVY)
    # Header row 3 for the native table.
    headers = ("Line item", "Q1", "Q2", "Q3", "Q4", "FY Total", "% of FY")
    for col, label in enumerate(headers, start=1):
        ws.cell(row=3, column=col, value=label).style = "BrandDocsHeader"
    # Demo body rows 4..7 with cross-quarter SUM + percent-of-total formulas.
    body = [
        ("Gross revenue", 320000, 351000, 372500, 410000),
        ("Discounts", -38400, -42120, -44700, -49200),
        ("Returns", -9600, -10530, -11175, -12300),
        ("Net revenue", None, None, None, None),  # formula row below
    ]
    for i, row in enumerate(body):
        r = 4 + i
        ws.cell(row=r, column=1, value=row[0])
        if row[0] == "Net revenue":
            # Net revenue = sum of the three lines above, per quarter.
            for col in range(2, 6):
                L = get_column_letter(col)
                c = ws.cell(row=r, column=col, value=f"=SUM({L}4:{L}6)")
                c.number_format = "#,##0"
        else:
            for col, val in zip(range(2, 6), row[1:]):
                c = ws.cell(row=r, column=col, value=val)
                c.number_format = "#,##0"
        # FY Total (col 6) = SUM of the four quarters in this row.
        ws.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})").number_format = "#,##0"
        # % of FY (col 7) = this row's FY total / net-revenue FY total (row 7).
        ws.cell(
            row=r, column=7, value=f"=IF($F$7=0,0,F{r}/$F$7)"
        ).number_format = "0.0%"
    # A grand-total SUBTOTAL row 9 (col 6).
    ws.cell(row=9, column=1, value="Subtotal (visible)")
    ws.cell(row=9, column=6, value="=SUBTOTAL(9,F4:F6)").number_format = "#,##0"
    # Native TABLE object over the body (header row 3 .. last data row 7).
    table = Table(displayName="BrandDocsDataTbl", ref="A3:G7")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True,
        showColumnStripes=False,
        showFirstColumn=False,
        showLastColumn=False,
    )
    ws.add_table(table)
    # CONDITIONAL FORMATTING: color scale on the quarter grid, a CellIs rule on
    # the % column, and a formula rule that flags negative FY totals.
    ws.conditional_formatting.add(
        "B4:E6",
        ColorScaleRule(
            start_type="min",
            start_color="FFF8696B",
            mid_type="percentile",
            mid_value=50,
            mid_color="FFFFEB84",
            end_type="max",
            end_color="FF63BE7B",
        ),
    )
    ws.conditional_formatting.add(
        "G4:G7",
        CellIsRule(
            operator="greaterThan",
            formula=["0.5"],
            fill=PatternFill("solid", fgColor=BRAND_AMBER),
        ),
    )
    ws.conditional_formatting.add(
        "F4:F6",
        FormulaRule(formula=["F4<0"], fill=PatternFill("solid", fgColor="FFFFC7CE")),
    )
    ws.freeze_panes = "B4"
    # Wider label column, tighter numeric columns so the table + chart fit a page.
    ws.column_dimensions["A"].width = 16
    for col in range(2, 8):
        ws.column_dimensions[get_column_letter(col)].width = 12

    # A native CHART driven by the model Net revenue row (one series, 4 points
    # across the Q1..Q4 categories).
    chart = BarChart()
    chart.title = "Quarterly net revenue"
    chart.type = "col"
    data = Reference(ws, min_col=2, min_row=7, max_col=5, max_row=7)  # B7:E7
    cats = Reference(ws, min_col=2, min_row=3, max_col=5, max_row=3)  # Q1..Q4
    chart.add_data(data, titles_from_data=False, from_rows=True)
    chart.set_categories(cats)
    # Brand: single navy/teal series, no redundant legend, clean axes.
    chart.series[0].graphicalProperties.solidFill = TEAL6
    chart.legend = None
    chart.y_axis.numFmt = "#,##0"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.width = 14
    chart.height = 8
    # Anchor the chart below the table so it stays within the print page.
    ws.add_chart(chart, "A11")
    # Contain the sheet on a single landscape page.
    ws.print_area = "A1:G27"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # A second chart (line) on the Summary sheet is added there.


def _build_summary(wb: Workbook) -> None:
    ws = wb.create_sheet("Summary")
    ws["A1"] = "Executive Summary"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color=BRAND_NAVY)
    ws["A3"] = "Metric"
    ws["B3"] = "Value"
    ws["A3"].style = "BrandDocsHeader"
    ws["B3"].style = "BrandDocsHeader"
    # CROSS-SHEET formulas: Summary pulls from Model and Inputs.
    ws["A4"] = "FY net revenue"
    ws["B4"] = "=Model!F7"
    ws["B4"].number_format = "#,##0"
    ws["A5"] = "Units sold (input)"
    ws["B5"] = "=Inputs!B4"
    ws["B5"].number_format = "#,##0"  # thousands separator: "29,070"
    ws["A6"] = "Unit price (input)"
    ws["B6"] = "=Inputs!B5"
    ws["B6"].style = "BrandDocsCurrency"
    ws["A7"] = "Implied gross"
    ws["B7"] = "=Inputs!B4*Inputs!B5"
    ws["B7"].style = "BrandDocsCurrency"
    ws["A8"] = "Net margin"
    ws["B8"] = "=IF(B7=0,0,B4/B7)"
    ws["B8"].style = "BrandDocsPercent"
    ws["A9"] = "Headline KPI"
    ws["B9"] = "85% net margin on $1.45M gross"  # single-cell named output slot
    ws.freeze_panes = "A4"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16

    # A LINE chart on the summary tracking Net revenue per quarter (cross-sheet
    # reference into the Model sheet), coherent with the Model bar chart.
    model = wb["Model"]
    chart = LineChart()
    chart.title = "Net revenue trend (Q1-Q4)"
    data = Reference(model, min_col=2, min_row=7, max_col=5, max_row=7)  # B7:E7
    cats = Reference(model, min_col=2, min_row=3, max_col=5, max_row=3)  # Q1..Q4
    chart.add_data(data, titles_from_data=False, from_rows=True)
    chart.set_categories(cats)
    # Name the single series so the legend reads sensibly.
    chart.series[0].tx = SeriesLabel(v="Net revenue")
    # Brand: amber line, clean axes, legend below to avoid extra width.
    chart.series[0].graphicalProperties.line.solidFill = AMBER6
    chart.series[0].graphicalProperties.line.width = 28000  # EMU (~2.2pt)
    chart.series[0].smooth = False
    chart.y_axis.numFmt = "#,##0"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = "b"
    chart.width = 14
    chart.height = 8
    # Anchor below the metric table, keep on a single page.
    ws.add_chart(chart, "A11")
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _build_dashboard(wb: Workbook) -> None:
    """Executive dashboard with formula-backed KPI cards and a native chart."""
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Executive Dashboard"
    ws["A1"].font = Font(name="Arial", size=18, bold=True, color=BRAND_NAVY)
    ws.merge_cells("A1:G1")
    ws["A2"] = (
        "Formula-backed snapshot generated from the model, inputs, and summary tabs"
    )
    ws["A2"].font = Font(name="Arial", size=11, italic=True, color=BRAND_TEAL)
    ws.merge_cells("A2:G2")

    kpis = [
        (1, "FY net revenue", "=Summary!B4", "#,##0"),
        (3, "Net margin", "=Summary!B8", "0.0%"),
        (5, "Input units", "=Summary!B5", "#,##0"),
        (7, "Scenario", "=Scenarios!B3", "@"),
    ]
    for start_col, label, formula, number_format in kpis:
        ws.merge_cells(
            start_row=4, start_column=start_col, end_row=4, end_column=start_col + 1
        )
        ws.merge_cells(
            start_row=5, start_column=start_col, end_row=5, end_column=start_col + 1
        )
        cell = ws.cell(row=4, column=start_col, value=label)
        cell.style = "BrandDocsHeader"
        value_cell = ws.cell(row=5, column=start_col, value=formula)
        value_cell.style = "BrandDocsKPI"
        value_cell.number_format = number_format
        value_cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    ws["A7"] = "Quarter"
    ws["B7"] = "Net revenue"
    ws["C7"] = "Growth"
    for cell in ws["A7:C7"][0]:
        cell.style = "BrandDocsHeader"
    for i, q in enumerate(("Q1", "Q2", "Q3", "Q4"), start=8):
        qidx = i - 7
        ws.cell(row=i, column=1, value=q)
        ws.cell(row=i, column=2, value=f"=Model!{get_column_letter(qidx + 1)}7")
        ws.cell(row=i, column=2).number_format = "#,##0"
        if i == 8:
            ws.cell(row=i, column=3, value="-")
        else:
            ws.cell(row=i, column=3, value=f"=IF(B{i - 1}=0,0,B{i}/B{i - 1}-1)")
            ws.cell(row=i, column=3).number_format = "0.0%"
    ws.conditional_formatting.add(
        "B8:B11",
        DataBarRule(start_type="min", end_type="max", color=TEAL6, showValue=True),
    )

    chart = BarChart()
    chart.title = "Net revenue"
    chart.add_data(
        Reference(ws, min_col=2, min_row=7, max_row=11), titles_from_data=True
    )
    chart.set_categories(Reference(ws, min_col=1, min_row=8, max_row=11))
    chart.series[0].graphicalProperties.solidFill = TEAL6
    chart.legend = None
    chart.width = 8.5
    chart.height = 6
    ws.add_chart(chart, "E8")
    for col, width in zip("ABCDEFGH", (12, 12, 12, 3, 12, 12, 12, 12)):
        ws.column_dimensions[col].width = width
    ws.print_area = "A1:H24"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _build_scenarios(wb: Workbook) -> None:
    ws = wb.create_sheet("Scenarios")
    ws["A1"] = "Scenario Controls"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color=BRAND_NAVY)
    ws["A3"] = "Selected scenario"
    ws["B3"] = "Base"
    ws["B3"].style = "BrandDocsInput"
    dv = DataValidation(
        type="list", formula1='"Base,Upside,Downside"', allow_blank=False
    )
    ws.add_data_validation(dv)
    dv.add(ws["B3"])
    ws["B3"].comment = Comment(
        "Synthetic scenario selector used by formulas and visual QA.", "BrandDocs"
    )
    headers = ("Scenario", "Revenue multiplier", "Margin delta", "Narrative")
    for col, label in enumerate(headers, start=1):
        ws.cell(row=5, column=col, value=label).style = "BrandDocsHeader"
    rows = [
        ("Base", 1.00, 0.000, "Plan case"),
        ("Upside", 1.12, 0.025, "Faster adoption"),
        ("Downside", 0.91, -0.035, "Cost pressure"),
    ]
    for r, row in enumerate(rows, start=6):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            if c in (2, 3):
                cell.style = "BrandDocsPercent" if c == 3 else "BrandDocsInput"
                cell.number_format = "0.0%" if c == 3 else "0.00x"
    ws["A10"] = "Scenario revenue"
    ws["B10"] = "=Summary!B4*INDEX(B6:B8,MATCH($B$3,A6:A8,0))"
    ws["B10"].number_format = "#,##0"
    ws["A11"] = "Scenario margin"
    ws["B11"] = "=Summary!B8+INDEX(C6:C8,MATCH($B$3,A6:A8,0))"
    ws["B11"].number_format = "0.0%"
    table = Table(displayName="BrandDocsScenarioTbl", ref="A5:D8")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    for col, width in zip("ABCD", (18, 18, 16, 24)):
        ws.column_dimensions[col].width = width
    ws.print_area = "A1:D13"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _build_data(wb: Workbook) -> None:
    """A raw transactions sheet (demo data the generator should clear/refill)."""
    ws = wb.create_sheet("Data")
    headers = ("Date", "Region", "Product", "Amount")
    for col, label in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=label).style = "BrandDocsHeader"
    demo = [
        ("2025-10-01", "North", "Widget", 12500),
        ("2025-10-02", "South", "Gadget", 9800),
        ("2025-10-03", "East", "Widget", 14200),
    ]
    for i, (d, region, prod, amt) in enumerate(demo):
        r = 2 + i
        dc = ws.cell(row=r, column=1, value=d)
        dc.number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=2, value=region)
        ws.cell(row=r, column=3, value=prod)
        ac = ws.cell(row=r, column=4, value=amt)
        ac.number_format = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'
    # A total row with a SUM the generator must preserve.
    ws.cell(row=5, column=3, value="Total")
    ws.cell(row=5, column=4, value="=SUM(D2:D4)").number_format = "#,##0"
    ws.freeze_panes = "A2"
    for col, w in zip("ABCD", (14, 12, 14, 16)):
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Named ranges (workbook scope) - the author's OWN vocabulary, surfaced as
# geometry by the extractor (never matched on as code literals).
# ---------------------------------------------------------------------------
def _add_named_ranges(wb: Workbook) -> None:
    defns = {
        # Single-cell title anchor under a merged header band.
        "report_title": "'Cover'!$A$1",
        "report_subtitle": "'Cover'!$A$2",
        "client_name": "'Cover'!$B$4",
        "period": "'Cover'!$B$5",
        # Multi-cell INPUTS block (sample-data candidate, in the frozen band edge).
        "inputs_block": "'Inputs'!$A$4:$C$7",
        # Multi-cell DATA region = body of the native table.
        "model_body": "'Model'!$A$4:$G$6",
        # A cross-sheet formula OUTPUT cell (single-cell named output slot).
        "headline_kpi": "'Summary'!$B$9",
        "dashboard_kpis": "'Dashboard'!$A$4:$H$5",
        "scenario_block": "'Scenarios'!$A$5:$D$8",
        # The raw demo-data block.
        "data_block": "'Data'!$A$2:$D$4",
    }
    for name in sorted(defns):
        wb.defined_names.add(DefinedName(name, attr_text=defns[name]))


_LOGO_CACHE: Path | None = None


def _logo_path() -> Path:
    """Materialize the shared BrandDocs wordmark to a temp file openpyxl can
    embed.

    Uses ``branddocs_mark_png`` so the cover logo is the SAME generated
    text-only wordmark used by the DOCX and PPTX examples.
    """
    global _LOGO_CACHE
    if _LOGO_CACHE is None:
        import tempfile

        tmp = Path(tempfile.gettempdir()) / "branddocs_template_wordmark.png"
        tmp.write_bytes(branddocs_mark_png(640, 160))
        _LOGO_CACHE = tmp
    return _LOGO_CACHE


def build(out: Path = OUT) -> Path:
    wb = Workbook()
    # Drop the default sheet; we author named sheets in a deliberate order.
    wb.remove(wb.active)
    _register_named_styles(wb)
    _build_cover(wb)
    _build_inputs(wb)
    _build_model(wb)
    _build_summary(wb)
    _build_dashboard(wb)
    _build_scenarios(wb)
    _build_data(wb)
    _add_named_ranges(wb)
    # Workbook-level: request a full recalc so authored formulas evaluate on open.
    wb.calculation.fullCalcOnLoad = True
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    freeze_ooxml(out)
    return out


if __name__ == "__main__":
    path = build()
    print(f"built {path}")
