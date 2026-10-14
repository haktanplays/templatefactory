"""
EXCEL BUILDER — Spec'ten satılabilir Excel çalışma kitabı üretir
================================================================
Sayfa düzeni: Start Here → Dashboard → veri sayfaları.
Her veri sayfası: başlık şeridi, stilli başlık satırı, örnek + boş satırlar,
dropdown doğrulama, sayı formatları, koşullu renkler, TOTAL satırı, grafik.
Dashboard: 4 KPI kartı (gerçek formüller) + 2 grafik + sayfa dizini.
"""
import logging
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from design_system import DesignSystem

logger = logging.getLogger("template-machine.excel")

HEADER_ROW = 5            # veri sayfalarında başlık satırı (1 başlık, 2 amaç, 3 nasıl kullanılır, 4 boşluk)
# Dashboard düzeni (qa.py de bu sabitleri kullanır)
DASH = {"kpi_row": 4, "metric_row": 8, "insight_row": 12, "inside_row": 20, "chart_anchors": ["F4", "F22"]}
DEFAULT_WIDTHS = {"text": 22, "currency": 14, "number": 12, "percent": 11, "date": 14, "dropdown": 16}
NUMERIC_TYPES = {"currency", "number", "percent"}


def _quote(sheet_name):
    return "'" + sheet_name.replace("'", "''") + "'"


def _parse_date(value):
    if isinstance(value, (datetime, date)):
        return value
    try:
        return datetime.strptime(str(value).strip()[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return value


def _number_format(col, symbol, sample_values=()):
    t = col["type"]
    if t == "currency":
        return f'"{symbol}"#,##0.00'
    if t == "percent":
        return "0%"
    if t == "date":
        return "d mmm yyyy"
    if t == "number":
        has_float = any(isinstance(v, float) and not float(v).is_integer() for v in sample_values)
        return "#,##0.00" if has_float else "#,##0"
    return "General"


class SheetLayout:
    """Dashboard formülleri ve grafikler için bir veri sayfasının koordinatları."""
    def __init__(self, ws, columns, data_start, sample_end, data_end):
        self.ws = ws
        self.name = ws.title
        self.columns = columns
        self.col_index = {c["name"]: i + 1 for i, c in enumerate(columns)}
        self.data_start = data_start      # ilk veri satırı
        self.sample_end = sample_end      # son örnek-veri satırı
        self.data_end = data_end          # son (boş dahil) veri satırı

    def letter(self, col_name):
        return get_column_letter(self.col_index[col_name])

    def range(self, col_name, samples_only=False):
        end = self.sample_end if samples_only else self.data_end
        return f"{_quote(self.name)}!${self.letter(col_name)}${self.data_start}:${self.letter(col_name)}${end}"


def build_workbook(spec, theme_name="ocean", brand="ProTemplates", fit_one_page=False):
    wb = Workbook()
    ds = DesignSystem(wb, theme=theme_name)
    ds.brand = brand
    excel = spec["excel"]
    symbol = excel.get("currency_symbol", "$")

    layouts = {}
    first = True
    for sh in excel["sheets"]:
        ws = wb.active if first else wb.create_sheet()
        first = False
        layouts[sh["name"]] = _build_data_sheet(ws, ds, sh, symbol, fit_one_page)

    dash = wb.create_sheet("Dashboard")
    _build_dashboard(dash, ds, spec, layouts, symbol, fit_one_page)
    start = wb.create_sheet("Start Here")
    _build_start_sheet(start, ds, spec, brand, fit_one_page)

    # Sıra: Start Here, Dashboard, veri sayfaları
    wb.move_sheet(start, offset=-(len(wb.sheetnames) - 1))
    wb.move_sheet(dash, offset=-(len(wb.sheetnames) - 2))
    wb.active = 1   # dosya Dashboard'da açılsın
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = (ws.title == "Dashboard")
    return wb


# ─── VERİ SAYFASI ───────────────────────────────────────
def _build_data_sheet(ws, ds, sh, symbol, fit_one_page):
    ws.title = sh["name"]
    ws.sheet_properties.tabColor = ds.t["secondary"]
    columns = sh["columns"]
    n = len(columns)
    last = get_column_letter(n)

    ds.style_dashboard_title(ws, sh["name"], sh.get("purpose", ""), f"A1:{last}1", row=1)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 18
    # Satır 3: nasıl kullanılır (açıklayıcı), satır 4: boşluk
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=n)
    how = ws.cell(row=3, column=1, value=("How to use: " + sh["how_to"]) if sh.get("how_to") else None)
    how.font = Font(name="Aptos", size=9, italic=True, color=ds.t["text_muted"])
    how.fill = ds.fill_bg_light
    how.alignment = Alignment(vertical="center", wrap_text=True)
    for c in range(1, n + 1):
        ws.cell(row=3, column=c).fill = ds.fill_bg_light
    ws.row_dimensions[3].height = 26 if sh.get("how_to") else 6
    ws.row_dimensions[4].height = 6

    # Başlık satırı (+ kolon yardım notları: hücre yorumu olarak)
    ds.style_header_row(ws, row=HEADER_ROW, col_start=1, col_end=n)
    for ci, col in enumerate(columns):
        cell = ws.cell(row=HEADER_ROW, column=ci + 1, value=col["name"])
        if col.get("help"):
            cm = Comment(col["help"], getattr(ds, "brand", "Template"))
            cm.width, cm.height = 260, 90
            cell.comment = cm

    data_cols = [c for c in columns if not c.get("formula")]
    rows = sh.get("rows", [])
    blank = int(sh.get("blank_rows", 10) or 0)
    data_start = HEADER_ROW + 1
    sample_end = HEADER_ROW + len(rows)
    data_end = sample_end + blank
    if data_end < data_start:
        data_end = data_start
    total_row = data_end + 1

    # Örnek veri
    for ri, row in enumerate(rows):
        r = data_start + ri
        for col, value in zip(data_cols, row):
            if col["type"] == "date":
                value = _parse_date(value)
            elif col["type"] in NUMERIC_TYPES and isinstance(value, str):
                try:
                    value = float(value.replace(",", "")) if value.strip() else None
                except ValueError:
                    pass
            if value == "":
                value = None
            ws.cell(row=r, column=columns.index(col) + 1, value=value)

    # Zebra + font + kenarlık
    ds.style_data_rows(ws, data_start, data_end, col_start=1, col_end=n)

    # Kolon bazlı format / hizalama / formül / doğrulama / koşullu renk
    samples_by_col = {c["name"]: [row[i] for row in rows] for i, c in enumerate(data_cols)}
    for ci, col in enumerate(columns):
        letter = get_column_letter(ci + 1)
        fmt = _number_format(col, symbol, samples_by_col.get(col["name"], ()))
        align = Alignment(horizontal="right" if col["type"] in NUMERIC_TYPES else
                          "center" if col["type"] in ("date", "dropdown") else "left",
                          vertical="center", wrap_text=col["type"] == "text")
        for r in range(data_start, data_end + 1):
            cell = ws.cell(row=r, column=ci + 1)
            cell.number_format = fmt
            cell.alignment = align

        if col.get("formula"):
            template = col["formula"].lstrip("=")
            for r in range(data_start, data_end + 1):
                f = template
                for other in columns:
                    f = f.replace("{" + other["name"] + "}", f"{get_column_letter(columns.index(other) + 1)}{r}")
                if r > sample_end:   # boş satırlarda ilk kolon doluncaya kadar boş görün
                    f = f'IF(A{r}="","",{f})'
                ws.cell(row=r, column=ci + 1, value="=" + f)

        if col["type"] == "dropdown":
            dv = DataValidation(type="list", formula1='"' + ",".join(col["options"]).replace('"', "") + '"',
                                allow_blank=True, showDropDown=False)
            dv.error = "Please pick a value from the list"
            dv.errorTitle = "Invalid entry"
            ws.add_data_validation(dv)
            dv.add(f"{letter}{data_start}:{letter}{data_end + 100}")
            for option, key in col.get("option_colors", {}).items():
                fill, font = ds.status_style(key)
                ws.conditional_formatting.add(
                    f"{letter}{data_start}:{letter}{data_end + 100}",
                    CellIsRule(operator="equal", formula=[f'"{option}"'], fill=fill, font=font))

        cond = col.get("conditional")
        rng = f"{letter}{data_start}:{letter}{data_end}"
        if cond == "negative_red":
            ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"],
                                                          font=ds.font_danger, fill=ds.fill_danger_light))
            ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"],
                                                          font=ds.font_success))
        elif cond == "data_bar":
            ds.add_data_bars(ws, letter, data_start, data_end)
        elif cond == "progress_bar":
            ws.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                           color=ds.t["success"], showValue=True))

        ws.column_dimensions[letter].width = col.get("width") or DEFAULT_WIDTHS.get(col["type"], 16)

    # TOTAL satırı
    totals = [c for c in columns if c.get("total")]
    if totals:
        ws.cell(row=total_row, column=1, value="TOTAL")
        for col in totals:
            ci = columns.index(col) + 1
            letter = get_column_letter(ci)
            func = {"sum": "SUM", "avg": "AVERAGE", "count": "COUNTA"}[col["total"]]
            c = ws.cell(row=total_row, column=ci, value=f"={func}({letter}{data_start}:{letter}{data_end})")
            c.number_format = _number_format(col, symbol, samples_by_col.get(col["name"], ()))
        ds.style_total_row(ws, total_row, col_start=1, col_end=n)
        for ci in range(2, n + 1):
            ws.cell(row=total_row, column=ci).alignment = Alignment(horizontal="right", vertical="center")

    layout = SheetLayout(ws, columns, data_start, sample_end, data_end)

    chart = sh.get("chart")
    if chart and rows:
        _add_chart(ws, ds, chart, layout, anchor=f"{get_column_letter(n + 2)}{HEADER_ROW}", width=16, height=9)
    if rows:
        _add_quick_stats(ws, ds, sh, layout, symbol, start_col=n + 2, start_row=HEADER_ROW + (14 if chart else 0))

    ds.finalize_sheet(ws, freeze=f"A{data_start}", gridlines=False)
    _print_setup(ws, fit_one_page)
    return layout


def _add_quick_stats(ws, ds, sh, layout, symbol, start_col, start_row):
    """Sayısal kolonlar için otomatik özet: ortalama / en yüksek / en düşük / en yüksek kalem."""
    numeric = [c for c in sh["columns"] if c["type"] in NUMERIC_TYPES or c.get("formula")][:3]
    if not numeric:
        return
    text_cols = [c for c in sh["columns"] if c["type"] == "text"]
    label_col = text_cols[0]["name"] if text_cols else None
    r0 = start_row
    ws.merge_cells(start_row=r0, start_column=start_col, end_row=r0, end_column=start_col + len(numeric))
    h = ws.cell(row=r0, column=start_col, value="QUICK STATS")
    h.font = ds.font_subheader
    ds.style_subheader_row(ws, r0 + 1, col_start=start_col, col_end=start_col + len(numeric))
    ws.cell(row=r0 + 1, column=start_col, value="")
    for i, c in enumerate(numeric):
        ws.cell(row=r0 + 1, column=start_col + 1 + i, value=c["name"])
    stats = [("Average", "AVERAGE"), ("Highest", "MAX"), ("Lowest", "MIN")]
    if label_col:
        stats.append(("Highest item", None))
    for si, (label, func) in enumerate(stats):
        r = r0 + 2 + si
        lc = ws.cell(row=r, column=start_col, value=label)
        lc.font = ds.font_data_bold
        lc.border = ds.border_light
        for i, c in enumerate(numeric):
            rng = f"{layout.letter(c['name'])}{layout.data_start}:{layout.letter(c['name'])}{layout.data_end}"
            if func:
                v = f"={func}({rng})"
                fmt = _number_format(c if not c.get("formula") else {"type": c.get("type", "number")}, symbol)
            else:
                lrng = f"{layout.letter(label_col)}{layout.data_start}:{layout.letter(label_col)}{layout.data_end}"
                v = f'=IFERROR(INDEX({lrng},MATCH(MAX({rng}),{rng},0)),"")'
                fmt = "General"
            cell = ws.cell(row=r, column=start_col + 1 + i, value=v)
            cell.number_format = fmt
            cell.font = ds.font_data
            cell.border = ds.border_light
            cell.alignment = Alignment(horizontal="right" if func else "center", vertical="center")
        ws.row_dimensions[r].height = 20
    for i in range(len(numeric) + 1):
        letter = get_column_letter(start_col + i)
        if (ws.column_dimensions[letter].width or 0) < 16:
            ws.column_dimensions[letter].width = 16


def _resolve_dashboard_formula(formula, kpi_cells, layouts):
    """{Label} -> dashboard hücresi, {Sheet!Column} -> veri aralığı."""
    def repl(m):
        ref = m.group(1)
        if ref in kpi_cells:
            return kpi_cells[ref]
        sheet, _, col = ref.partition("!")
        return layouts[sheet].range(col)
    import re
    return re.sub(r"\{([^}]+)\}", repl, formula)


def _add_chart(ws, ds, chart_cfg, layout, anchor, width=16, height=8.5):
    kind = chart_cfg["type"]
    chart = {"bar": BarChart, "line": LineChart, "pie": PieChart}[kind]()
    chart.title = chart_cfg.get("title") or None
    chart.width, chart.height = width, height
    src = layout.ws
    cats = Reference(src, min_col=layout.col_index[chart_cfg["category_col"]],
                     min_row=layout.data_start, max_row=layout.sample_end)
    for vc in chart_cfg["value_cols"]:
        ci = layout.col_index[vc]
        data = Reference(src, min_col=ci, min_row=layout.data_start - 1, max_row=layout.sample_end)
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    palette = [ds.t["primary"], ds.t["secondary"], ds.t["accent"], ds.t["danger"]]
    if kind == "pie":
        from openpyxl.chart.series import DataPoint
        series = chart.series[0]
        for i in range(layout.sample_end - layout.data_start + 1):
            pt = DataPoint(idx=i)
            pt.graphicalProperties.solidFill = palette[i % len(palette)] if i < 4 else _tint(palette[i % len(palette)], i)
            series.dPt.append(pt)
        chart.legend.position = "r"
    else:
        for i, s in enumerate(chart.series):
            s.graphicalProperties.solidFill = palette[i % len(palette)]
            s.graphicalProperties.line.solidFill = palette[i % len(palette)]
        chart.y_axis.majorGridlines = None
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.legend.position = "b"
        if kind == "bar":
            chart.type = "col"
            chart.gapWidth = 60
    ws.add_chart(chart, anchor)


def _tint(hex_color, i):
    # aynı paleti biraz açarak daha fazla pasta dilimi için renk üret
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    k = 0.25 + 0.1 * (i % 4)
    return "%02X%02X%02X" % tuple(int(c + (255 - c) * k) for c in (r, g, b))


# ─── DASHBOARD ──────────────────────────────────────────
def _build_dashboard(ws, ds, spec, layouts, symbol, fit_one_page):
    ws.sheet_properties.tabColor = ds.t["primary"]
    for col in "ABCD":
        ws.column_dimensions[col].width = 24
    ws.column_dimensions["E"].width = 3
    for col in "FGHI":
        ws.column_dimensions[col].width = 24

    ds.style_dashboard_title(ws, spec["product_name"], spec.get("tagline", ""), "A1:I1", row=1)
    ws.row_dimensions[1].height = 34
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 10
    explain_font = Font(name="Aptos", size=9, italic=True, color=ds.t["text_muted"])
    explain_align = Alignment(wrap_text=True, vertical="top", horizontal="center")
    dash = spec["excel"]["dashboard"]
    kpi_cells = {}

    def card_row(items, row, build_formula):
        for ki, item in enumerate(items[:4]):
            col = ki + 1
            formula = build_formula(item)
            fmt = {"currency": f'"{symbol}"#,##0.00', "percent": "0%"}.get(item.get("format"), "#,##0")
            ds.style_kpi_card(ws, row=row, col=col, label=item["label"].upper(), value_cell_ref=formula, number_format=fmt)
            kpi_cells[item["label"]] = f"${get_column_letter(col)}${row + 1}"
            ex = ws.cell(row=row + 2, column=col, value=item.get("explain") or None)
            ex.font = explain_font
            ex.alignment = explain_align
            ex.fill = ds.fill_white
            ex.border = ds.border_light
        ws.row_dimensions[row].height = 20
        ws.row_dimensions[row + 1].height = 40
        ws.row_dimensions[row + 2].height = 48
        ws.row_dimensions[row + 3].height = 8

    def kpi_formula(kpi):
        lay = layouts[kpi["sheet"]]
        rng = lay.range(kpi["column"])
        if kpi["agg"] == "sumif":
            return f'=SUMIF({lay.range(kpi["where_column"])},"{kpi.get("equals", "")}",{rng})'
        func = {"sum": "SUM", "avg": "AVERAGE", "count": "COUNTA", "max": "MAX", "min": "MIN"}[kpi["agg"]]
        return f"={func}({rng})"

    card_row(dash.get("kpis", []), DASH["kpi_row"], kpi_formula)
    card_row(dash.get("metrics", []), DASH["metric_row"],
             lambda m: _resolve_dashboard_formula(m["formula"], kpi_cells, layouts))

    # Key insights: formülle yazılan cümleler
    r = DASH["insight_row"]
    ws.cell(row=r, column=1, value="KEY INSIGHTS").font = ds.font_subheader
    for ins in dash.get("insights", [])[:5]:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        value = _resolve_dashboard_formula(ins, kpi_cells, layouts) if ins.startswith("=") else ins
        c = ws.cell(row=r, column=1, value=value)
        c.font = ds.font_data
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.fill = ds.fill_bg_light
        for cc in range(1, 5):
            ws.cell(row=r, column=cc).fill = ds.fill_bg_light
            ws.cell(row=r, column=cc).border = ds.border_light
        ws.row_dimensions[r].height = 30

    charts = dash.get("charts", [])
    for ch, anchor in zip(charts[:2], DASH["chart_anchors"]):
        _add_chart(ws, ds, ch, layouts[ch["sheet"]], anchor=anchor, width=15.5, height=8.5)

    # Sayfa dizini
    r = max(DASH["inside_row"], r + 2)
    ws.cell(row=r, column=1, value="WHAT'S INSIDE").font = ds.font_subheader
    r += 1
    for sh in spec["excel"]["sheets"]:
        ws.cell(row=r, column=1, value=sh["name"]).font = ds.font_data_bold
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        d = ws.cell(row=r, column=2, value=sh.get("purpose", ""))
        d.font = ds.font_muted
        d.alignment = Alignment(vertical="center", wrap_text=True)
        r += 1

    ds.finalize_sheet(ws, freeze=None, gridlines=False)
    _print_setup(ws, fit_one_page)


# ─── START HERE ─────────────────────────────────────────
def _build_start_sheet(ws, ds, spec, brand, fit_one_page):
    ws.sheet_properties.tabColor = ds.t["accent"]
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 110
    ds.style_dashboard_title(ws, spec["product_name"], spec.get("tagline", ""), "A1:B1", row=1)
    ws.row_dimensions[1].height = 34
    ws.row_dimensions[2].height = 20

    r = 4
    ws.cell(row=r, column=2, value="HOW TO USE THIS TEMPLATE").font = ds.font_subheader
    r += 1
    for i, step in enumerate(spec.get("instructions", []), 1):
        ws.cell(row=r, column=1, value=f"{i}.").font = ds.font_data_bold
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="right", vertical="top")
        c = ws.cell(row=r, column=2, value=step)
        c.font = ds.font_data
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 15 * max(1, (len(step) // 95) + 1)
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="WHAT'S INSIDE").font = ds.font_subheader
    r += 1
    ws.cell(row=r, column=2, value="Dashboard — totals and charts that update automatically from the other tabs").font = ds.font_data
    r += 1
    for sh in spec["excel"]["sheets"]:
        ws.cell(row=r, column=2, value=f"{sh['name']} — {sh.get('purpose', '')}").font = ds.font_data
        r += 1

    r += 1
    tips = [
        "Cells with a small arrow are dropdown lists: pick a value instead of typing it so colours and totals stay accurate.",
        "Rows highlighted in colour are driven by their Status column. Blank rows at the bottom of every table are ready for your own data.",
        "Do not delete the TOTAL row or the Dashboard formulas; just overwrite the sample data.",
        "Google Sheets: File > Import > Upload this file. Formulas, dropdowns and colours carry over.",
    ]
    ws.cell(row=r, column=2, value="TIPS").font = ds.font_subheader
    r += 1
    for tip in tips:
        c = ws.cell(row=r, column=2, value="•  " + tip)
        c.font = ds.font_data
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 15 * max(1, (len(tip) // 95) + 1)
        r += 1

    r += 1
    ws.cell(row=r, column=2, value=f"© {datetime.now().year} {brand}. For personal use. Thank you for your purchase!").font = ds.font_muted
    ds.finalize_sheet(ws, freeze=None, gridlines=False)
    _print_setup(ws, fit_one_page, landscape=False)


def _print_setup(ws, fit_one_page, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if fit_one_page else 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
