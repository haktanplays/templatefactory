# Template Machine — Automated Digital Product Factory

## Description
An autonomous pipeline that discovers trending keywords, generates premium digital products (Excel workbooks + printable PDF planners + listing images), publishes them to Gumroad, exports Etsy CSVs and reports results.

Quality model: a **product spec** (JSON) describes every sheet, column type, sample row, formula, chart and PDF page. Claude writes the spec for the keyword; `excel_builder.py` and `pdf_builder.py` render it. Hand-crafted specs in `category_packs/` are the quality bar (given to Claude as the example) and the fallback when no API key is set.

## Supported categories (v2)
`budget`, `wedding`, `meal_prep` — a keyword must match one of these (see `spec.py` → `CATEGORY_KEYWORDS`). Add a category by dropping a new `category_packs/<name>.json` and its keywords into `spec.py`.

## Trigger
- **Cron**: Every day at 08:00 UTC → run `main.py`
- **Manual**: User says "generate templates", "run template machine", or "find trends"
- **Keyword**: User provides a specific keyword → skip trend discovery, go straight to production

## Workflow
1. **Discover**: Scrape Etsy autocomplete + Google Trends for keywords in the supported categories
2. **Score**: Rank opportunities by demand, specificity and template type
3. **Generate**: keyword → spec (Claude or category pack) → Excel (Start Here, Dashboard with KPIs + charts, 4 data sheets with dropdowns, totals, conditional colours) + PDF planner (cover, tables, checklists, weekly grid, habit tracker, notes) + 2000×1600 listing image built from real screenshots (needs LibreOffice + pdftoppm; falls back to a drawn mock)
4. **Publish**: Gumroad upload (token required) + Etsy bulk CSV
5. **Report**: Daily summary in the log and `run_history.json`

## Commands
- `python main.py` — full pipeline
- `python main.py --trends-only` — discover and report trends only
- `python main.py --keyword "wedding guest list tracker"` — generate one product set
- `python main.py --status` — stats from past runs

## Files
- `main.py` — pipeline orchestrator
- `spec.py` — spec schema, validation, category detection, pack loading
- `category_packs/*.json` — hand-made specs per category
- `product_generator.py` — Claude spec generation + orchestration of the builders
- `excel_builder.py` — spec → styled .xlsx (uses `design_system.py`)
- `pdf_builder.py` — spec → printable A4 PDF
- `preview_generator.py` — listing image from real workbook screenshots
- `trend_scraper.py`, `gumroad_publisher.py`, `etsy_exporter.py` — discovery and distribution
- `config.json` — model, pricing, themes, trend seeds; `.env` — API keys

## Requirements
- Python 3.10+, `pip install -r requirements.txt`
- Optional but recommended: LibreOffice (`soffice`) and poppler (`pdftoppm`) for real-screenshot listing images
- API keys in `.env`: `ANTHROPIC_API_KEY` (spec generation; without it the category pack is used as-is), `GUMROAD_ACCESS_TOKEN`

## Setup
```bash
pip install -r requirements.txt
cp .env.example .env        # then add your keys
python main.py --keyword "monthly budget planner"
```
