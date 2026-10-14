# Template Machine — Automated Digital Product Factory

## Description
An autonomous pipeline that discovers trending keywords, generates premium digital products (Excel workbooks + printable PDF planners + listing images), runs every product through an automated quality gate, publishes the ones that pass to Gumroad, exports Etsy CSVs and reports results.

Quality model: a **product spec** (JSON) describes every sheet, column type, sample row, formula, chart, dashboard metric, insight sentence, explanation and PDF page. Claude writes the spec for the keyword; `excel_builder.py` and `pdf_builder.py` render it. Hand-crafted specs in `category_packs/` are the quality bar (given to Claude as the example) and the fallback when no API key is set or Claude's spec fails QA.

## Supported categories
`budget`, `wedding`, `meal_prep` — a keyword must match one of these (see `spec.py` → `CATEGORY_KEYWORDS`). Add a category by dropping a new `category_packs/<name>.json` and its keywords into `spec.py`, then run `python main.py --selftest`.

## What a product contains
- **Excel**: Start Here (how-to, tips), Dashboard (4 KPIs + 4 derived metrics, each with a plain-language explanation; 3-5 insight sentences written by formulas so they change with the data; 2 charts), 4 data sheets with a "how to use" line, header tooltips, dropdowns with status colours, number formats, conditional formatting, TOTAL rows, a Quick Stats box (average / highest / lowest / top item) and a chart.
- **PDF** (A4): cover, "how to read your numbers" guide, tables with prefilled labels, checklists, weekly grid, habit tracker, boxes, notes.
- **Listing image** 2000×1600 built from real workbook screenshots (needs LibreOffice + pdftoppm; falls back to a drawn mock).
- **Listing text**: Etsy-ready title (≤140), description, 13 tags.

## Quality gate (`qa.py`)
Runs on every product before publishing. Excel is recalculated with LibreOffice and scanned for `#VALUE!`/`#REF!`/`#NAME?`/`#DIV/0!`; KPI/metric cells must be numbers, insights must be sentences, TOTAL rows must equal the Python sum of the sample data, dropdown lists must stay under Excel's 255-character limit, sheet names/merges/freeze panes/charts are checked, placeholder text ("Item A", "lorem", "TBD") is rejected. PDF: page count, titles rendered, no blank pages. Image: size and real screenshot present. Listing: Etsy limits.
Flow: Claude spec fails → regenerate with the failure list (`qa.max_claude_attempts`) → fall back to the category pack → still failing → product is not published and the run report says why. `qa.strict` (config.json) makes a missing LibreOffice a failure; without LibreOffice formulas cannot be verified.

## Trigger
- **Cron**: Every day at 08:00 UTC → run `main.py`
- **Manual**: User says "generate templates", "run template machine", or "find trends"
- **Keyword**: User provides a specific keyword → skip trend discovery, go straight to production

## Commands
- `python main.py` — full pipeline
- `python main.py --trends-only` — discover and report trends only
- `python main.py --keyword "wedding guest list tracker"` — generate one product set
- `python main.py --selftest` — build every category pack in every theme and run the QA gate (exit code 1 on failure)
- `python main.py --status` — stats from past runs

## Files
- `main.py` — pipeline orchestrator
- `spec.py` — spec schema, validation, category detection, pack loading
- `category_packs/*.json` — hand-made specs per category
- `product_generator.py` — Claude spec generation, build orchestration, QA retry loop
- `excel_builder.py` — spec → styled .xlsx (uses `design_system.py`)
- `pdf_builder.py` — spec → printable A4 PDF
- `preview_generator.py` — listing image from real workbook screenshots
- `qa.py` — quality gate
- `trend_scraper.py`, `gumroad_publisher.py`, `etsy_exporter.py` — discovery and distribution
- `config.json` — model, pricing, themes, trend seeds, QA settings; `.env` — API keys

## Requirements
- Python 3.10+, `pip install -r requirements.txt`
- **LibreOffice** (`soffice`) and poppler (`pdftoppm`): required for the formula check in the quality gate and for real-screenshot listing images. Debian/Ubuntu: `apt install libreoffice-calc poppler-utils`; macOS: `brew install --cask libreoffice && brew install poppler`
- API keys in `.env`: `ANTHROPIC_API_KEY` (spec generation; without it the category pack is used as-is), `GUMROAD_ACCESS_TOKEN`

## Setup
```bash
pip install -r requirements.txt
cp .env.example .env        # then add your keys
python main.py --selftest   # should print 12/12
python main.py --keyword "monthly budget planner"
```
