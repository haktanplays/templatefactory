"""
PRODUCT SPEC — Şema, doğrulama ve kategori paketleri
=====================================================
Bir "spec" ürünün tam planıdır: Excel sayfaları (kolon tipleri, örnek veri,
formüller, grafikler), PDF sayfaları ve mağaza listesi metni.
Claude bu spec'i üretir; excel_builder / pdf_builder sadece çizer.
category_packs/*.json dosyaları elle hazırlanmış, satılabilir kalitede
örneklerdir: Claude'a şablon olarak verilir, API yoksa doğrudan kullanılır.
"""
import json, re, logging
from pathlib import Path

logger = logging.getLogger("template-machine.spec")

PACK_DIR = Path(__file__).parent / "category_packs"

# keyword içindeki kelimelere göre kategori tespiti (ilk eşleşen kazanır)
CATEGORY_KEYWORDS = {
    "wedding":   ["wedding", "bride", "bridal", "groom", "marriage", "engagement", "honeymoon"],
    "meal_prep": ["meal", "recipe", "grocery", "food", "diet", "nutrition", "macro", "calorie", "cooking", "menu"],
    "budget":    ["budget", "expense", "finance", "money", "saving", "savings", "debt", "spending", "bill", "income", "cash"],
}

COLUMN_TYPES = {"text", "currency", "number", "percent", "date", "dropdown"}
TOTAL_AGGS = {"sum", "avg", "count"}
CONDITIONALS = {"negative_red", "data_bar", "progress_bar"}
KPI_AGGS = {"sum", "avg", "count", "max", "min", "sumif"}
CHART_TYPES = {"bar", "pie", "line"}
PDF_PAGE_TYPES = {"cover", "guide", "table", "checklist", "weekly_grid", "habit_tracker", "boxes", "notes"}
METRIC_FORMATS = {"currency", "number", "percent"}
OPTION_COLOR_KEYS = {"success", "warning", "danger", "info"}
RESERVED_SHEET_NAMES = {"dashboard", "start here"}
DROPDOWN_FORMULA_LIMIT = 255   # Excel: liste doğrulaması 255 karakteri aşarsa dosya "onarım" hatası verir


class SpecError(ValueError):
    pass


def detect_category(keyword):
    """Keyword'den kategori döndürür; desteklenmeyen keyword için None."""
    kw = keyword.lower()
    for cat, words in CATEGORY_KEYWORDS.items():
        if any(w in kw for w in words):
            return cat
    return None


def supported_categories():
    return sorted(p.stem for p in PACK_DIR.glob("*.json"))


def load_pack(category):
    path = PACK_DIR / f"{category}.json"
    if not path.exists():
        raise SpecError(f"Kategori paketi yok: {category} (mevcut: {supported_categories()})")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ─── DOĞRULAMA ──────────────────────────────────────────
def _req(d, key, where):
    if key not in d:
        raise SpecError(f"{where}: '{key}' eksik")
    return d[key]


def validate_spec(spec):
    """Spec'i doğrular, eksik alanları varsayılanlarla doldurur ve döndürür.
    Claude'dan gelen JSON'un çiziciyi çökertmemesi için tek kapı burasıdır."""
    if not isinstance(spec, dict):
        raise SpecError("spec bir JSON nesnesi olmalı")
    _req(spec, "product_name", "spec")
    spec.setdefault("tagline", "")
    spec.setdefault("instructions", [])
    spec.setdefault("theme", "ocean")

    excel = _req(spec, "excel", "spec")
    excel.setdefault("currency_symbol", "$")
    sheets = _req(excel, "sheets", "excel")
    if not sheets:
        raise SpecError("excel.sheets boş")

    seen_names = set()
    for si, sh in enumerate(sheets):
        where = f"sheets[{si}]"
        name = str(_req(sh, "name", where))
        # Excel sayfa adı kuralları: 31 karakter, \ / ? * [ ] : yok
        name = re.sub(r'[\\/?*\[\]:]', '-', name).strip()[:31] or f"Sheet{si+1}"
        if name.lower() in seen_names or name.lower() in RESERVED_SHEET_NAMES:
            name = f"{name[:28]} {si+1}"
        seen_names.add(name.lower())
        sh["name"] = name
        sh.setdefault("purpose", "")
        sh["how_to"] = str(sh.get("how_to") or "")
        sh.setdefault("blank_rows", 10)
        cols = _req(sh, "columns", where)
        if not cols:
            raise SpecError(f"{where}: columns boş")
        col_names = []
        for ci, col in enumerate(cols):
            cwhere = f"{where}.columns[{ci}]"
            if isinstance(col, str):            # kısa yazım: sadece ad
                col = {"name": col, "type": "text"}
                cols[ci] = col
            cname = str(_req(col, "name", cwhere)).strip()
            if not cname:
                raise SpecError(f"{cwhere}: boş kolon adı")
            col["name"] = cname
            col["help"] = str(col.get("help") or "")
            col_names.append(cname)
            ctype = col.get("type", "text")
            if ctype not in COLUMN_TYPES:
                col["type"] = "text"
            if col.get("type") == "dropdown":
                opts = []
                for o in col.get("options", []):
                    o = str(o).replace(",", " /").replace('"', "").strip()
                    if o and o not in opts:
                        opts.append(o)
                while opts and len(",".join(opts)) + 2 > DROPDOWN_FORMULA_LIMIT:
                    opts.pop()
                if not opts:
                    col["type"] = "text"
                col["options"] = opts
                oc = col.get("option_colors") or {}
                col["option_colors"] = {k.replace(",", " /"): v for k, v in oc.items()}
            if col.get("total") not in TOTAL_AGGS:
                col.pop("total", None)
            if col.get("conditional") not in CONDITIONALS:
                col.pop("conditional", None)
            oc = col.get("option_colors") or {}
            col["option_colors"] = {k: v for k, v in oc.items() if v in OPTION_COLOR_KEYS}
        if len(set(col_names)) != len(col_names):
            raise SpecError(f"{where}: kolon adları tekrar ediyor")

        # formül kolonları: {Kolon Adı} yer tutucuları gerçek kolonlara işaret etmeli
        col_types = {c["name"]: ("number" if c.get("formula") else c["type"]) for c in cols}
        for col in cols:
            if col.get("formula"):
                f = col["formula"]
                for ref in re.findall(r"\{([^}]+)\}", f):
                    if ref not in col_names:
                        raise SpecError(f"{where}: formül bilinmeyen kolona başvuruyor: {ref}")
                    # metin/dropdown kolonu aritmetiğe giremez; sadece IF(...="...") karşılaştırmasında kullanılabilir
                    if col_types[ref] in ("text", "dropdown", "date") and not re.search(r"\{" + re.escape(ref) + r"\}\s*(=|<>)", f):
                        raise SpecError(f"{where}: formül sayısal olmayan kolonla hesap yapıyor: {ref}")
                if col["type"] not in ("currency", "number", "percent"):
                    col["type"] = "number"

        # örnek satırlar: formül kolonları hariç kolon sayısı kadar değer
        data_cols = [c for c in cols if not c.get("formula")]
        rows = sh.get("rows") or []
        fixed = []
        for ri, row in enumerate(rows):
            if not isinstance(row, (list, tuple)):
                continue
            row = list(row)[:len(data_cols)]
            row += [""] * (len(data_cols) - len(row))
            for i, c in enumerate(data_cols):
                row[i] = _coerce_value(c, row[i])
            fixed.append(row)
        sh["rows"] = fixed
        # yüzde kolonu: değerlerin çoğu 1'den büyükse 0-100 ölçeğinde yazılmış demektir
        for i, c in enumerate(data_cols):
            if c["type"] == "percent":
                nums = [r[i] for r in fixed if isinstance(r[i], (int, float))]
                if nums and sum(1 for v in nums if v > 1) > len(nums) / 2:
                    for r in fixed:
                        if isinstance(r[i], (int, float)):
                            r[i] = r[i] / 100

        chart = sh.get("chart")
        if chart and not _valid_chart(chart, col_names):
            sh.pop("chart", None)

    dash = excel.setdefault("dashboard", {})
    sheet_cols = {sh["name"]: [c["name"] for c in sh["columns"]] for sh in sheets}
    kpis = []
    for k in dash.get("kpis", []):
        if k.get("sheet") in sheet_cols and k.get("column") in sheet_cols[k["sheet"]] and k.get("agg", "sum") in KPI_AGGS:
            k.setdefault("agg", "sum")
            k.setdefault("format", "number")
            if k["agg"] == "sumif" and k.get("where_column") not in sheet_cols[k["sheet"]]:
                continue
            kpis.append(k)
    for k in kpis:
        k["explain"] = str(k.get("explain") or "")
    dash["kpis"] = kpis[:4]

    # Türetilmiş metrikler ve içgörü cümleleri: {KPI Label} veya {Sheet!Column} yer tutucuları
    labels = [k["label"] for k in dash["kpis"]]
    metrics = []
    for m in dash.get("metrics", []):
        if not isinstance(m, dict) or not m.get("label") or not str(m.get("formula", "")).startswith("="):
            continue
        bad = _bad_placeholders(m["formula"], labels, sheet_cols)
        if bad:
            logger.warning(f"metrik '{m['label']}' atlandı, çözülemeyen başvuru: {bad}")
            continue
        m["format"] = m.get("format") if m.get("format") in METRIC_FORMATS else "number"
        m["explain"] = str(m.get("explain") or "")
        metrics.append(m)
        labels.append(m["label"])
    dash["metrics"] = metrics[:4]
    insights = []
    for ins in dash.get("insights", []):
        if isinstance(ins, dict):
            ins = ins.get("formula") or ins.get("text") or ""
        ins = str(ins).strip()
        if not ins:
            continue
        if ins.startswith("="):
            bad = _bad_placeholders(ins, labels, sheet_cols)
            if bad:
                logger.warning(f"insight atlandı, çözülemeyen başvuru: {bad}")
                continue
        insights.append(ins)
    dash["insights"] = insights[:5]

    charts = []
    for ch in dash.get("charts", []):
        if ch.get("sheet") in sheet_cols and _valid_chart(ch, sheet_cols[ch["sheet"]]):
            charts.append(ch)
    dash["charts"] = charts[:2]

    pdf = spec.setdefault("pdf", {})
    pdf.setdefault("title", spec["product_name"])
    pages = [p for p in pdf.get("pages", []) if isinstance(p, dict) and p.get("type") in PDF_PAGE_TYPES]
    for p in pages:
        if p["type"] == "guide":
            p["items"] = [{"heading": str(i.get("heading", "")), "text": str(i.get("text", ""))}
                          for i in p.get("items", []) if isinstance(i, dict) and i.get("text")][:8]
    pages = [p for p in pages if p["type"] != "guide" or p["items"]]
    if not pages or pages[0].get("type") != "cover":
        pages.insert(0, {"type": "cover"})
    if pages[-1].get("type") != "notes":
        pages.append({"type": "notes", "title": "Notes"})
    pdf["pages"] = pages

    listing = spec.setdefault("listing", {})
    listing.setdefault("title", spec["product_name"])
    listing["title"] = str(listing["title"])[:140]
    listing.setdefault("description", spec.get("tagline", ""))
    listing["tags"] = [str(t)[:20] for t in listing.get("tags", [])][:13]
    return spec


def _coerce_value(col, value):
    """Örnek hücre değerini kolon tipine uydurur; uymayan sayısal değer boş kalır (SUM'ı bozmasın)."""
    if value is None or value == "":
        return ""
    t = col["type"]
    if t in ("currency", "number", "percent"):
        if isinstance(value, bool):
            return ""
        if isinstance(value, (int, float)):
            return value
        try:
            return float(str(value).replace(",", "").replace("$", "").replace("%", "").strip())
        except ValueError:
            return ""
    if t == "date":
        m = re.match(r"\d{4}-\d{2}-\d{2}", str(value))
        return m.group(0) if m else ""
    if t == "dropdown":
        v = str(value).strip()
        return v if v in col["options"] else ""
    return str(value)


def _bad_placeholders(formula, labels, sheet_cols):
    """Dashboard formülündeki {X} başvurularından çözülemeyenleri döndürür."""
    bad = []
    for ref in re.findall(r"\{([^}]+)\}", formula):
        if ref in labels:
            continue
        if "!" in ref:
            sheet, _, col = ref.partition("!")
            if sheet in sheet_cols and col in sheet_cols[sheet]:
                continue
        bad.append(ref)
    return bad


def _valid_chart(chart, col_names):
    if chart.get("type") not in CHART_TYPES:
        return False
    if chart.get("category_col") not in col_names:
        return False
    vals = [v for v in chart.get("value_cols", []) if v in col_names]
    if not vals:
        return False
    chart["value_cols"] = vals[:3]
    chart.setdefault("title", "")
    return True


# ─── CLAUDE PROMPT İÇİN ŞEMA AÇIKLAMASI ────────────────
SCHEMA_DOC = """
JSON SPEC FORMAT (follow exactly; the example below is a complete, valid spec):

{
  "category": "<category id>",
  "theme": "ocean" | "sunset" | "forest" | "midnight",
  "product_name": "short, specific product name",
  "tagline": "one line benefit statement",
  "target_audience": "who buys this",
  "instructions": ["4-5 short 'how to use' steps, one per tab"],
  "excel": {
    "currency_symbol": "$",
    "sheets": [
      {
        "name": "max 31 chars, no / \\\\ ? * [ ] :",
        "purpose": "one sentence",
        "how_to": "one sentence telling the user how to work with this tab",
        "columns": [
          {"name": "Column", "type": "text|currency|number|percent|date|dropdown",
           "help": "one sentence shown as a header tooltip explaining what to enter",
           "options": ["only for dropdown"],
           "option_colors": {"Option": "success|warning|danger|info"},
           "total": "sum|avg|count",                 // optional, adds a TOTAL row
           "conditional": "negative_red|data_bar|progress_bar",   // optional
           "formula": "={Other Column}-{Another Column}",          // optional computed column; reference columns by {name}
           "width": 18}
        ],
        "rows": [["one value per NON-formula column, in column order", "...", 123.45, "2026-03-01"]],
        "blank_rows": 10,
        "chart": {"type": "bar|pie|line", "title": "...", "category_col": "Column", "value_cols": ["Column"]}
      }
    ],
    "dashboard": {
      "kpis": [{"label": "...", "sheet": "Sheet Name", "column": "Column", "agg": "sum|avg|count|max|min|sumif", "where_column": "Column", "equals": "Value", "format": "currency|number|percent",
                "explain": "one sentence: what this number means and a sensible target"}],
      "metrics": [{"label": "Savings Rate", "formula": "=IF({Total Income}=0,0,({Total Income}-{Total Spent})/{Total Income})", "format": "percent",
                   "explain": "..."}],
      "insights": ["=\"Biggest overspend: \"&INDEX({Sheet Name!Category},MATCH(MIN({Sheet Name!Difference}),{Sheet Name!Difference},0))",
                   "=\"Guests still pending: \"&COUNTIF({Guest List!RSVP},\"Pending\")"],
      "charts": [{"type": "bar|pie|line", "title": "...", "sheet": "Sheet Name", "category_col": "Column", "value_cols": ["Column"]}]
    }
  },
  "pdf": {
    "title": "...",
    "pages": [
      {"type": "cover"},
      {"type": "guide", "title": "How to Use This Planner", "items": [{"heading": "...", "text": "2-3 sentences"}]},
      {"type": "table", "title": "...", "subtitle": "...", "columns": ["A", "B"], "col_widths": [2, 1], "prefill": ["optional first-column labels"], "rows": 20, "total_row": true},
      {"type": "checklist", "title": "...", "sections": {"Section": ["item", "item"]}, "columns": 1},
      {"type": "weekly_grid", "title": "...", "columns": ["Breakfast", "Lunch", "Dinner"]},
      {"type": "habit_tracker", "title": "...", "habits": ["habit", "habit"]},
      {"type": "boxes", "title": "...", "boxes": ["box label", "box label"], "box_height": 42},
      {"type": "notes", "title": "Notes"}
    ]
  },
  "listing": {"title": "Etsy SEO title, max 140 chars", "description": "150-250 words", "tags": ["13 tags, each max 20 chars"]}
}

RULES:
- Dates are ISO strings "YYYY-MM-DD". Money and numbers are plain numbers, not strings.
- 4 sheets, 10-18 realistic sample rows each (7 for a weekly plan). Sample data must be specific to the keyword's domain: real-sounding names, items, amounts and dates. Never generic placeholders like "Item A".
- Every sheet needs at least one numeric column; at least two sheets need a "total" column and a dropdown column with option_colors.
- Dashboard: exactly 4 KPIs (each with "explain"), 4 derived "metrics" and 3-5 "insights", 2 charts. In metric/insight formulas reference
  KPI or metric labels as {Label} and sheet columns as {Sheet Name!Column}; use only INDEX, MATCH, MIN, MAX, SUM, SUMIF, COUNTIF, COUNTA,
  AVERAGE, IF, TEXT and & so the formulas work in Excel, Google Sheets and LibreOffice. Insights must read as full sentences with the number inside.
- Every sheet has "how_to"; every column has "help". Explanations are written for a non-expert: what the number means, why it matters, what a good value is.
- PDF: 6-8 pages using the page types above, each one useful on paper for this specific keyword (prefill table labels where sensible).
"""
