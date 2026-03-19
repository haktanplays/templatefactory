# Template Machine — Automated Digital Product Factory

## Description
An autonomous pipeline that discovers trending keywords, generates premium digital products (Excel templates + printable PDFs), publishes them to Gumroad, and reports results via messaging.

## Trigger
- **Cron**: Every day at 08:00 UTC → run `main.py`
- **Manual**: User says "generate templates", "run template machine", or "find trends"
- **Keyword**: User provides a specific keyword → skip trend discovery, go straight to production

## Workflow
1. **Discover**: Scrape Google Trends + Etsy autocomplete for trending template keywords
2. **Score**: Filter and rank opportunities by demand, competition, and price potential
3. **Generate**: For each top keyword, produce Excel template + PDF planner using Claude API
4. **Publish**: Upload to Gumroad with optimized titles, descriptions, tags, and pricing
5. **Report**: Send summary to user via configured messaging channel

## Commands
- `template machine run` — Execute full pipeline
- `template machine trends` — Only discover and report trends (no generation)
- `template machine generate <keyword>` — Generate products for a specific keyword
- `template machine status` — Show stats (total products, revenue, top sellers)

## Files
- `main.py` — Pipeline orchestrator
- `trend_scraper.py` — Trend discovery engine
- `product_generator.py` — Excel + PDF generation with Claude API
- `gumroad_publisher.py` — Gumroad API integration
- `design_system.py` — Premium Excel styling engine
- `config.json` — API keys and preferences

## Requirements
- Python 3.10+
- pip: openpyxl, reportlab, pytrends, requests, anthropic
- API Keys: Anthropic (Claude), Gumroad

## Setup
```bash
cd ~/.openclaw/skills/template-machine
pip install openpyxl reportlab pytrends requests anthropic
cp config.example.json config.json
# Edit config.json with your API keys
```
