"""
PDF BUILDER — Spec'ten yazdırılabilir A4 planner üretir (reportlab)
===================================================================
Sayfa tipleri: cover, table, checklist, weekly_grid, habit_tracker, boxes, notes.
Her sayfa tema rengiyle başlık şeridi, alt bilgide marka + sayfa numarası taşır.
"""
import logging
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table,
                                TableStyle, PageBreak, KeepTogether)

logger = logging.getLogger("template-machine.pdf")

THEME_COLORS = {
    "ocean":    {"accent": "#1B4965", "dark": "#0B2545", "light": "#EEF4F8", "border": "#C9D6E3", "muted": "#6B7B8D", "soft": "#5FA8D3"},
    "sunset":   {"accent": "#6C2E7A", "dark": "#2D1B3D", "light": "#FDF5F0", "border": "#E3CFDB", "muted": "#8B7B8D", "soft": "#D4638A"},
    "forest":   {"accent": "#1A472A", "dark": "#0D2818", "light": "#F0F7F2", "border": "#C3D8C6", "muted": "#6B8B6D", "soft": "#4A9A5B"},
    "midnight": {"accent": "#1A1A2E", "dark": "#0A0A1A", "light": "#F2F2F8", "border": "#CFCFDF", "muted": "#6B6B8D", "soft": "#E94560"},
}

PAGE_W, PAGE_H = A4
MARGIN = 15 * mm
CONTENT_W = PAGE_W - 2 * MARGIN
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


class _Theme:
    def __init__(self, name):
        tc = THEME_COLORS.get(name, THEME_COLORS["ocean"])
        self.accent = colors.HexColor(tc["accent"])
        self.dark = colors.HexColor(tc["dark"])
        self.light = colors.HexColor(tc["light"])
        self.border = colors.HexColor(tc["border"])
        self.muted = colors.HexColor(tc["muted"])
        self.soft = colors.HexColor(tc["soft"])
        self.white = colors.white
        self.h1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=30, leading=36, textColor=self.dark, alignment=TA_LEFT)
        self.h2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=self.accent)
        self.sub = ParagraphStyle("sub", fontName="Helvetica", fontSize=10.5, leading=14, textColor=self.muted)
        self.body = ParagraphStyle("body", fontName="Helvetica", fontSize=10, leading=14, textColor=self.dark)
        self.label = ParagraphStyle("label", fontName="Helvetica-Bold", fontSize=9, leading=11, textColor=self.accent)
        self.cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=9, leading=11, textColor=self.dark)
        self.head = ParagraphStyle("head", fontName="Helvetica-Bold", fontSize=9, leading=11, textColor=self.white, alignment=TA_CENTER)
        self.item = ParagraphStyle("item", fontName="Helvetica", fontSize=10, leading=13, textColor=self.dark)
        self.tiny = ParagraphStyle("tiny", fontName="Helvetica", fontSize=7, leading=8, textColor=self.muted, alignment=TA_CENTER)


def build_pdf(spec, theme_name, path, brand="ProTemplates", tagline=""):
    t = _Theme(theme_name)
    pdf = spec["pdf"]
    title = pdf.get("title", spec["product_name"])
    footer_text = f"{title}  •  © {datetime.now().year} {brand}"

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(t.accent)
        canvas.rect(0, PAGE_H - 6 * mm, PAGE_W, 6 * mm, stroke=0, fill=1)
        canvas.setFillColor(t.muted)
        canvas.setFont("Helvetica", 8)
        canvas.drawString(MARGIN, 8 * mm, footer_text)
        canvas.drawRightString(PAGE_W - MARGIN, 8 * mm, f"Page {doc.page}")
        canvas.restoreState()

    def on_cover(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(t.dark)
        canvas.rect(0, PAGE_H * 0.62, PAGE_W, PAGE_H * 0.38, stroke=0, fill=1)
        canvas.setFillColor(t.soft)
        canvas.rect(0, PAGE_H * 0.62 - 3 * mm, PAGE_W, 3 * mm, stroke=0, fill=1)
        canvas.setFillColor(t.white)
        canvas.setFont("Helvetica-Bold", 11)
        canvas.drawString(MARGIN, PAGE_H - 20 * mm, brand.upper())
        canvas.setFont("Helvetica-Bold", 30)
        y = PAGE_H * 0.62 + 40 * mm
        for line in _wrap(title, 26):
            canvas.drawString(MARGIN, y, line)
            y -= 13 * mm
        canvas.setFont("Helvetica", 12)
        canvas.setFillColor(colors.HexColor("#DDE3EA"))
        canvas.drawString(MARGIN, PAGE_H * 0.62 + 14 * mm, spec.get("tagline", ""))
        canvas.setFillColor(t.muted)
        canvas.setFont("Helvetica", 8)
        canvas.drawString(MARGIN, 8 * mm, f"© {datetime.now().year} {brand}  •  {tagline}")
        canvas.restoreState()

    frame = Frame(MARGIN, 14 * mm, CONTENT_W, PAGE_H - 14 * mm - 12 * mm, id="main", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    cover_frame = Frame(MARGIN, 14 * mm, CONTENT_W, PAGE_H * 0.62 - 20 * mm, id="cover", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc = BaseDocTemplate(str(path), pagesize=A4, title=title, author=brand,
                          leftMargin=MARGIN, rightMargin=MARGIN, topMargin=12 * mm, bottomMargin=14 * mm)
    doc.addPageTemplates([PageTemplate(id="cover", frames=[cover_frame], onPage=on_cover),
                          PageTemplate(id="main", frames=[frame], onPage=on_page)])

    story = []
    renderers = {"cover": _cover, "guide": _guide, "table": _table, "checklist": _checklist, "weekly_grid": _weekly_grid,
                 "habit_tracker": _habit_tracker, "boxes": _boxes, "notes": _notes}
    pages = pdf["pages"]
    for i, page in enumerate(pages):
        if page["type"] == "cover":
            story += _cover(page, spec, t)
            story.append(_NextTemplate("main"))
        else:
            story += renderers[page["type"]](page, spec, t)
        if i < len(pages) - 1:
            story.append(PageBreak())
    doc.build(story)
    logger.info(f"📄 PDF oluşturuldu: {path} ({len(pages)} sayfa tipi, tema: {theme_name})")
    return str(path)


from reportlab.platypus.doctemplate import NextPageTemplate as _NextTemplate


def _wrap(text, width):
    words, lines, cur = text.split(), [], ""
    for w in words:
        if len(cur) + len(w) + 1 > width and cur:
            lines.append(cur)
            cur = w
        else:
            cur = (cur + " " + w).strip()
    if cur:
        lines.append(cur)
    return lines[:3]


def _heading(page, t, default=""):
    out = [Paragraph(page.get("title", default), t.h2)]
    rule = Table([[""]], colWidths=[CONTENT_W], rowHeights=[1.2 * mm])
    rule.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), t.accent)]))
    out.append(rule)
    if page.get("subtitle"):
        out.append(Spacer(1, 3))
        out.append(Paragraph(page["subtitle"], t.sub))
    out.append(Spacer(1, 6 * mm))
    return out


# ─── SAYFA TİPLERİ ──────────────────────────────────────
def _cover(page, spec, t):
    pages = [p for p in spec["pdf"]["pages"] if p["type"] not in ("cover",)]
    items = [Paragraph("WHAT'S INSIDE", t.label), Spacer(1, 3)]
    for p in pages:
        items.append(Paragraph(f"•  {p.get('title', p['type'].replace('_', ' ').title())}", t.body))
    items.append(Spacer(1, 8 * mm))
    items.append(Paragraph("HOW TO USE", t.label))
    items.append(Spacer(1, 3))
    for line in ["Print on A4 (or Letter, fit to page). Single-sided works best for the trackers.",
                 "Pair it with the Excel workbook: plan on paper, track on screen.",
                 "Print the tracker pages as many times as you need."]:
        items.append(Paragraph(f"•  {line}", t.body))
    return [Spacer(1, 6 * mm)] + items


def _guide(page, spec, t):
    """Açıklayıcı rehber: her metrik/bölüm için başlık + 2-3 cümle."""
    out = _heading(page, t, "How to Use This Planner")
    for item in page.get("items", []):
        cell = [[Paragraph(item["heading"].upper(), t.label)], [Paragraph(item["text"], t.body)]]
        tbl = Table(cell, colWidths=[CONTENT_W])
        tbl.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), t.light), ("LINEBELOW", (0, 0), (0, 0), 0.6, t.accent),
                                 ("BOX", (0, 0), (-1, -1), 0.6, t.border), ("LEFTPADDING", (0, 0), (-1, -1), 6),
                                 ("RIGHTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 1), (-1, 1), 5), ("BOTTOMPADDING", (0, 1), (-1, 1), 7)]))
        out += [KeepTogether(tbl), Spacer(1, 4 * mm)]
    return out


def _table(page, spec, t):
    cols = page.get("columns", ["Item", "Details"])
    widths = page.get("col_widths") or [1] * len(cols)
    widths = widths[:len(cols)] + [1] * (len(cols) - len(widths))
    total = sum(widths)
    col_w = [CONTENT_W * w / total for w in widths]
    prefill = page.get("prefill", [])
    n_rows = max(int(page.get("rows", 20)), len(prefill))
    data = [[Paragraph(c, t.head) for c in cols]]
    for i in range(n_rows):
        first = prefill[i] if i < len(prefill) else ""
        data.append([Paragraph(first, t.cell)] + [""] * (len(cols) - 1))
    if page.get("total_row"):
        data.append([Paragraph("TOTAL", t.head)] + [""] * (len(cols) - 1))
    row_h = min(9 * mm, (PAGE_H - 70 * mm) / (len(data)))
    tbl = Table(data, colWidths=col_w, rowHeights=[8 * mm] + [row_h] * (len(data) - 1))
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), t.accent),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, -1), 0.5, t.border),
        ("LINEBEFORE", (1, 1), (-1, -1), 0.4, t.border),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [t.white, t.light]),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    if page.get("total_row"):
        style += [("BACKGROUND", (0, -1), (-1, -1), t.accent)]
    tbl.setStyle(TableStyle(style))
    return _heading(page, t, "Tracker") + [tbl]


def _checklist(page, spec, t):
    sections = page.get("sections") or {"Checklist": [""] * 12}
    ncols = 2 if page.get("columns") == 2 or len(sections) >= 5 else 1
    blocks = []
    for name, items in sections.items():
        rows = [[Paragraph(name, t.label), ""]]
        for it in items:
            rows.append(["", Paragraph(it, t.item)])
        for _ in range(2 if items and any(items) else 0):
            rows.append(["", ""])       # kendi maddelerin için boş satır
        box = 4.2 * mm
        w = CONTENT_W / ncols - (4 * mm if ncols == 2 else 0)
        tbl = Table(rows, colWidths=[7 * mm, w - 7 * mm], rowHeights=[6 * mm] + [7 * mm] * (len(rows) - 1))
        st = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("SPAN", (0, 0), (1, 0)),
              ("LINEBELOW", (0, 0), (-1, 0), 0.8, t.accent),
              ("LINEBELOW", (1, 1), (1, -1), 0.4, t.border),
              ("LEFTPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
        for r in range(1, len(rows)):
            st.append(("BOX", (0, r), (0, r), 0.6, t.accent))   # tik kutusu
        tbl.setStyle(TableStyle(st))
        # kutunun kare görünmesi için hücre içini daraltan iç tablo yerine BOX hücreye uygulanır; genişlik 7mm, yükseklik 7mm
        blocks.append([tbl, Spacer(1, 4 * mm)])
    if ncols == 1:
        return _heading(page, t, "Checklist") + [KeepTogether(b) for b in blocks]
    left = [f for b in blocks[0::2] for f in b]
    right = [f for b in blocks[1::2] for f in b]
    grid = Table([[left, right]], colWidths=[CONTENT_W / 2, CONTENT_W / 2])
    grid.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (0, 0), 4 * mm), ("LEFTPADDING", (1, 0), (1, 0), 4 * mm)]))
    return _heading(page, t, "Checklist") + [grid]


def _weekly_grid(page, spec, t):
    cols = page.get("columns") or ["Breakfast", "Lunch", "Dinner"]
    cols = cols[:5]
    day_w = 26 * mm
    col_w = [day_w] + [(CONTENT_W - day_w) / len(cols)] * len(cols)
    data = [[Paragraph("", t.head)] + [Paragraph(c, t.head) for c in cols]]
    for d in DAYS:
        data.append([Paragraph(d, t.label)] + [""] * len(cols))
    row_h = (PAGE_H - 75 * mm) / 7
    tbl = Table(data, colWidths=col_w, rowHeights=[8 * mm] + [row_h] * 7)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), t.accent),
        ("BACKGROUND", (0, 1), (0, -1), t.light),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, t.border),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]))
    return _heading(page, t, "Weekly Plan") + [tbl]


def _habit_tracker(page, spec, t):
    habits = (page.get("habits") or ["Habit"])[:8]
    days = 31
    label_w = 44 * mm
    cell_w = (CONTENT_W - label_w) / days
    header = [Paragraph("Habit", t.head)] + [Paragraph(str(d), ParagraphStyle("d", parent=t.head, fontSize=6.5)) for d in range(1, days + 1)]
    data = [header]
    for h in habits:
        data.append([Paragraph(h, t.cell)] + [""] * days)
    tbl = Table(data, colWidths=[label_w] + [cell_w] * days, rowHeights=[7 * mm] + [9 * mm] * len(habits))
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), t.accent),
        ("GRID", (0, 0), (-1, -1), 0.4, t.border),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
    ]))
    out = _heading(page, t, "Habit Tracker") + [Paragraph("Month: ____________________", t.sub), Spacer(1, 4 * mm), tbl]
    # ikinci bir blok: aylık özet / ödül
    out += [Spacer(1, 10 * mm), Paragraph("THIS MONTH'S FOCUS", t.label), Spacer(1, 2 * mm)]
    box = Table([[""]], colWidths=[CONTENT_W], rowHeights=[40 * mm])
    box.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, t.border), ("BACKGROUND", (0, 0), (-1, -1), t.light)]))
    out += [box, Spacer(1, 8 * mm), Paragraph("REWARD WHEN I HIT 20+ DAYS", t.label), Spacer(1, 2 * mm)]
    box2 = Table([[""]], colWidths=[CONTENT_W], rowHeights=[22 * mm])
    box2.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, t.border)]))
    out.append(box2)
    return out


def _boxes(page, spec, t):
    labels = (page.get("boxes") or ["Notes"])[:5]
    h = float(page.get("box_height", 40)) * mm
    out = _heading(page, t, "Planner")
    for label in labels:
        tbl = Table([[Paragraph(label, t.label)], [""]], colWidths=[CONTENT_W], rowHeights=[7 * mm, h])
        tbl.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), t.light), ("BOX", (0, 0), (-1, -1), 0.8, t.border),
                                 ("LINEBELOW", (0, 0), (0, 0), 0.6, t.accent), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                                 ("LEFTPADDING", (0, 0), (-1, -1), 5)]))
        out += [tbl, Spacer(1, 5 * mm)]
    return out


def _notes(page, spec, t):
    n = 26
    rows = [[""] for _ in range(n)]
    tbl = Table(rows, colWidths=[CONTENT_W], rowHeights=[(PAGE_H - 70 * mm) / n] * n)
    tbl.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.5, t.border)]))
    return _heading(page, t, "Notes") + [tbl]
