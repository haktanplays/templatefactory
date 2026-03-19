"""
EXCEL DESIGN SYSTEM — Premium Template Styling Engine
======================================================
Bu modül her otomatik üretilen template'e profesyonel,
tutarlı, satılabilir kalitede görsellik kazandırır.

Kullanım:
    from excel_design_system import DesignSystem
    ds = DesignSystem(wb, theme="ocean")  # veya "sunset", "forest", "midnight"
    ds.style_header_row(ws, row=1, cols=6)
    ds.style_data_range(ws, start=2, end=20, cols=6)
    ds.add_dashboard_title(ws, "My Dashboard", "A1:F1")
"""

from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill, NamedStyle
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, DataBarRule, ColorScaleRule
from copy import copy

THEMES = {
    "ocean": {
        "name": "Ocean",
        "primary": "1B4965",
        "primary_light": "2D6A8F",
        "secondary": "5FA8D3",
        "accent": "62B6CB",
        "accent_warm": "BEE9E8",
        "bg_dark": "0B2545",
        "bg_medium": "13315C",
        "bg_light": "EEF4F8",
        "bg_alt_row": "F5F9FC",
        "text_light": "FFFFFF",
        "text_dark": "1B2A4A",
        "text_muted": "6B7B8D",
        "success": "2ECC71",
        "warning": "F39C12",
        "danger": "E74C3C",
        "border": "D1DCE8",
        "border_light": "E8EFF5",
    },
    "sunset": {
        "name": "Sunset",
        "primary": "6C2E7A",
        "primary_light": "8E4497",
        "secondary": "D4638A",
        "accent": "F7A072",
        "accent_warm": "FFECD2",
        "bg_dark": "2D1B3D",
        "bg_medium": "4A2C5E",
        "bg_light": "FDF5F0",
        "bg_alt_row": "FDF8F5",
        "text_light": "FFFFFF",
        "text_dark": "2D1B3D",
        "text_muted": "8B7B8D",
        "success": "27AE60",
        "warning": "E67E22",
        "danger": "C0392B",
        "border": "E8D5E0",
        "border_light": "F2E8EE",
    },
    "forest": {
        "name": "Forest",
        "primary": "1A472A",
        "primary_light": "2D6A3E",
        "secondary": "4A9A5B",
        "accent": "8FD694",
        "accent_warm": "D4EDDA",
        "bg_dark": "0D2818",
        "bg_medium": "1A472A",
        "bg_light": "F0F7F2",
        "bg_alt_row": "F5FAF6",
        "text_light": "FFFFFF",
        "text_dark": "1A2E1A",
        "text_muted": "6B8B6D",
        "success": "28A745",
        "warning": "FFC107",
        "danger": "DC3545",
        "border": "C8DCC8",
        "border_light": "DEE8DE",
    },
    "midnight": {
        "name": "Midnight",
        "primary": "1A1A2E",
        "primary_light": "16213E",
        "secondary": "0F3460",
        "accent": "E94560",
        "accent_warm": "FFE0E6",
        "bg_dark": "0A0A1A",
        "bg_medium": "1A1A2E",
        "bg_light": "F2F2F8",
        "bg_alt_row": "F7F7FB",
        "text_light": "FFFFFF",
        "text_dark": "1A1A2E",
        "text_muted": "6B6B8D",
        "success": "00D2FF",
        "warning": "FFB347",
        "danger": "E94560",
        "border": "D0D0E0",
        "border_light": "E8E8F0",
    },
}


class DesignSystem:
    def __init__(self, workbook, theme="ocean"):
        self.wb = workbook
        self.t = THEMES.get(theme, THEMES["ocean"])
        self.theme_name = theme
        self._setup_fonts()
        self._setup_fills()
        self._setup_borders()
        self._setup_alignments()

    def _setup_fonts(self):
        self.font_title = Font(name="Aptos", size=18, bold=True, color=self.t["text_light"])
        self.font_subtitle = Font(name="Aptos", size=11, color=self.t["accent"])
        self.font_header = Font(name="Aptos", size=10, bold=True, color=self.t["text_light"])
        self.font_subheader = Font(name="Aptos", size=9, bold=True, color=self.t["primary"])
        self.font_data = Font(name="Aptos", size=10, color=self.t["text_dark"])
        self.font_data_bold = Font(name="Aptos", size=10, bold=True, color=self.t["text_dark"])
        self.font_muted = Font(name="Aptos", size=9, color=self.t["text_muted"])
        self.font_total = Font(name="Aptos", size=11, bold=True, color=self.t["text_light"])
        self.font_success = Font(name="Aptos", size=10, bold=True, color=self.t["success"])
        self.font_danger = Font(name="Aptos", size=10, bold=True, color=self.t["danger"])
        self.font_kpi_value = Font(name="Aptos", size=22, bold=True, color=self.t["primary"])
        self.font_kpi_label = Font(name="Aptos", size=9, color=self.t["text_muted"])
        self.font_kpi_delta = Font(name="Aptos", size=9, bold=True)

    def _setup_fills(self):
        self.fill_primary = PatternFill("solid", fgColor=self.t["primary"])
        self.fill_primary_light = PatternFill("solid", fgColor=self.t["primary_light"])
        self.fill_secondary = PatternFill("solid", fgColor=self.t["secondary"])
        self.fill_accent = PatternFill("solid", fgColor=self.t["accent"])
        self.fill_bg_dark = PatternFill("solid", fgColor=self.t["bg_dark"])
        self.fill_bg_light = PatternFill("solid", fgColor=self.t["bg_light"])
        self.fill_bg_alt = PatternFill("solid", fgColor=self.t["bg_alt_row"])
        self.fill_white = PatternFill("solid", fgColor="FFFFFF")
        self.fill_success_light = PatternFill("solid", fgColor="E8F8EE")
        self.fill_danger_light = PatternFill("solid", fgColor="FDECEE")
        self.fill_warning_light = PatternFill("solid", fgColor="FFF8E8")
        self.fill_accent_warm = PatternFill("solid", fgColor=self.t["accent_warm"])

    def _setup_borders(self):
        thin = Side(style="thin", color=self.t["border"])
        thin_light = Side(style="thin", color=self.t["border_light"])
        medium = Side(style="medium", color=self.t["primary"])
        self.border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.border_light = Border(left=thin_light, right=thin_light, top=thin_light, bottom=thin_light)
        self.border_bottom = Border(bottom=thin)
        self.border_bottom_thick = Border(bottom=medium)
        self.border_none = Border()

    def _setup_alignments(self):
        self.align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        self.align_right = Alignment(horizontal="right", vertical="center")
        self.align_title = Alignment(horizontal="left", vertical="center")

    # ─── HIGH-LEVEL STYLING METHODS ────────────────────────────

    def style_dashboard_title(self, ws, title, subtitle, merge_range, row=1):
        ws.merge_cells(merge_range)
        cell = ws.cell(row=row, column=1)
        cell.value = title
        cell.font = self.font_title
        cell.fill = self.fill_bg_dark
        cell.alignment = self.align_title

        from openpyxl.utils import column_index_from_string
        end_part = merge_range.split(":")[-1]
        end_col_letter = ''.join(c for c in end_part if c.isalpha())
        start_part = merge_range.split(":")[0]
        start_col_letter = ''.join(c for c in start_part if c.isalpha())
        start_col_num = column_index_from_string(start_col_letter)
        end_col_num = column_index_from_string(end_col_letter)

        for c in range(start_col_num, end_col_num + 1):
            ws.cell(row=row, column=c).fill = self.fill_bg_dark

        if subtitle:
            ws.merge_cells(f"A{row+1}:{end_col_letter}{row+1}")
            sub_cell = ws.cell(row=row + 1, column=1)
            sub_cell.value = subtitle
            sub_cell.font = self.font_subtitle
            sub_cell.fill = self.fill_bg_dark
            sub_cell.alignment = self.align_title
            for c in range(start_col_num, end_col_num + 1):
                ws.cell(row=row + 1, column=c).fill = self.fill_bg_dark

    def style_header_row(self, ws, row, col_start=1, col_end=10):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = self.font_header
            cell.fill = self.fill_primary
            cell.alignment = self.align_center
            cell.border = self.border_all
        ws.row_dimensions[row].height = 32

    def style_subheader_row(self, ws, row, col_start=1, col_end=10):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = self.font_subheader
            cell.fill = self.fill_bg_light
            cell.alignment = self.align_center
            cell.border = self.border_all
        ws.row_dimensions[row].height = 26

    def style_data_rows(self, ws, start_row, end_row, col_start=1, col_end=10, zebra=True):
        for r in range(start_row, end_row + 1):
            is_alt = (r - start_row) % 2 == 1
            for c in range(col_start, col_end + 1):
                cell = ws.cell(row=r, column=c)
                cell.font = self.font_data
                cell.fill = self.fill_bg_alt if (zebra and is_alt) else self.fill_white
                cell.alignment = self.align_center if c > 1 else self.align_left
                cell.border = self.border_light
            ws.row_dimensions[r].height = 24

    def style_total_row(self, ws, row, col_start=1, col_end=10):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = self.font_total
            cell.fill = self.fill_primary
            cell.alignment = self.align_center
            cell.border = self.border_all
        ws.row_dimensions[row].height = 30

    def style_kpi_card(self, ws, row, col, label, value_cell_ref, unit="", delta=None):
        label_cell = ws.cell(row=row, column=col)
        label_cell.value = label
        label_cell.font = self.font_kpi_label
        label_cell.fill = self.fill_white
        label_cell.alignment = Alignment(horizontal="center", vertical="bottom")
        label_cell.border = self.border_light

        val_cell = ws.cell(row=row + 1, column=col)
        val_cell.value = value_cell_ref if isinstance(value_cell_ref, str) and value_cell_ref.startswith("=") else value_cell_ref
        val_cell.font = self.font_kpi_value
        val_cell.fill = self.fill_white
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.border = self.border_light

        if delta is not None:
            delta_cell = ws.cell(row=row + 2, column=col)
            delta_cell.value = delta
            delta_cell.font = Font(name="Aptos", size=9, bold=True, color=self.t["success"] if "+" in str(delta) or (isinstance(delta, (int, float)) and delta > 0) else self.t["danger"])
            delta_cell.fill = self.fill_white
            delta_cell.alignment = Alignment(horizontal="center", vertical="top")
            delta_cell.border = self.border_light

        for r in range(row, row + (3 if delta else 2)):
            ws.cell(row=r, column=col).border = self.border_light

    def add_data_bars(self, ws, col_letter, start_row, end_row):
        rule = DataBarRule(
            start_type="min", end_type="max",
            color=self.t["secondary"],
            showValue=True, minLength=5, maxLength=90
        )
        ws.conditional_formatting.add(f"{col_letter}{start_row}:{col_letter}{end_row}", rule)

    def add_color_scale(self, ws, col_letter, start_row, end_row, reverse=False):
        if reverse:
            rule = ColorScaleRule(
                start_type="min", start_color=self.t["success"],
                mid_type="percentile", mid_value=50, mid_color="FFFFFF",
                end_type="max", end_color=self.t["danger"]
            )
        else:
            rule = ColorScaleRule(
                start_type="min", start_color=self.t["danger"],
                mid_type="percentile", mid_value=50, mid_color="FFFFFF",
                end_type="max", end_color=self.t["success"]
            )
        ws.conditional_formatting.add(f"{col_letter}{start_row}:{col_letter}{end_row}", rule)

    def add_conditional_positive_negative(self, ws, col_letter, start_row, end_row):
        ws.conditional_formatting.add(
            f"{col_letter}{start_row}:{col_letter}{end_row}",
            CellIsRule(operator="greaterThan", formula=["0"],
                       font=self.font_success, fill=self.fill_success_light)
        )
        ws.conditional_formatting.add(
            f"{col_letter}{start_row}:{col_letter}{end_row}",
            CellIsRule(operator="lessThan", formula=["0"],
                       font=self.font_danger, fill=self.fill_danger_light)
        )

    def auto_column_widths(self, ws, col_widths):
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

    def add_chart(self, ws, chart_type, data_range, cats_range, title, position, size=(14, 10)):
        if chart_type == "bar":
            chart = BarChart()
            chart.style = 26
        elif chart_type == "line":
            chart = LineChart()
            chart.style = 26
        elif chart_type == "pie":
            chart = PieChart()

        chart.title = title
        chart.width = size[0]
        chart.height = size[1]

        data = Reference(ws, **data_range)
        cats = Reference(ws, **cats_range)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        if chart_type != "pie":
            chart.y_axis.title = None
            chart.x_axis.title = None
            chart.legend.position = "b"

        for i, series in enumerate(chart.series):
            colors = [self.t["primary"], self.t["secondary"], self.t["accent"], self.t["danger"]]
            series.graphicalProperties.solidFill = colors[i % len(colors)]

        ws.add_chart(chart, position)

    def freeze_panes(self, ws, cell="A4"):
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = cell

    def set_print_area(self, ws, area):
        ws.print_area = area
        ws.sheet_properties.pageSetUpPr = None

    def finalize_sheet(self, ws, freeze="A4", gridlines=False):
        ws.sheet_view.showGridLines = gridlines
        if freeze:
            ws.freeze_panes = freeze


# ─── CONVENIENCE FUNCTION ────────────────────────────────

def create_styled_workbook(theme="ocean"):
    from openpyxl import Workbook
    wb = Workbook()
    ds = DesignSystem(wb, theme=theme)
    return wb, ds
