"""
ETSY EXPORTER — Generates CSV files for Etsy bulk upload
Etsy Seller Handbook: Use "Add listings with a CSV" feature
"""
import csv, logging, re
from pathlib import Path
from datetime import datetime

logger = logging.getLogger("template-machine.etsy-export")


def export_single(product_data, config, writer):
    """Tek bir ürün setini CSV satırlarına yazar."""
    listing = product_data.get("listing", {})
    files = product_data.get("files", {})
    pricing = product_data.get("pricing", {})
    keyword = product_data.get("keyword", "")
    concept = product_data.get("concept", {})

    brand = config["design"]["brand_name"]
    base_title = listing.get("title", f"{keyword.title()} Template")
    description = listing.get("description", f"Professional {keyword} template")
    tags = listing.get("tags", [])

    # Etsy max 13 tags
    etsy_tags = tags[:13]
    # Pad with generic tags if less than 13
    generic_tags = ["template", "spreadsheet", "planner", "digital download",
                    "printable", "excel template", "instant download", "tracker",
                    "organizer", "productivity", "business template", "planning tool", "digital planner"]
    for gt in generic_tags:
        if len(etsy_tags) >= 13:
            break
        if gt not in etsy_tags:
            etsy_tags.append(gt)

    tags_str = ",".join(etsy_tags[:13])

    # Excel listing
    if files.get("excel"):
        excel_desc = f"""{description}

WHAT YOU GET:
- 1x Excel Template (.xlsx) - Works with Microsoft Excel and Google Sheets
- Professional dashboard with automated calculations
- Ready-to-use formulas (SUM, AVERAGE, conditional formatting)
- Clean, modern design with {product_data.get('theme', 'ocean').title()} theme

HOW TO USE:
1. Download the file after purchase
2. Open in Excel or upload to Google Sheets
3. Replace sample data with your own
4. Formulas update automatically!

INSTANT DOWNLOAD - No physical item will be shipped.

© {datetime.now().year} {brand}"""

        writer.writerow({
            "title": f"{base_title} Excel Spreadsheet Digital Download",
            "description": excel_desc,
            "price": f"{pricing.get('excel', 14.99):.2f}",
            "currency": "USD",
            "tags": tags_str,
            "materials": "",
            "quantity": "999",
            "sku": f"TM-EXCEL-{re.sub(r'[^A-Z0-9]', '', keyword.upper()[:20])}",
            "file": Path(files["excel"]).name if files.get("excel") else "",
            "preview_image": Path(files["preview"]).name if files.get("preview") else "",
            "category": "Craft Supplies & Tools > Templates",
            "type": "digital",
            "who_made": "i_did",
            "when_made": "2020_2025",
            "is_supply": "true",
        })

    # PDF listing
    if files.get("pdf"):
        pdf_desc = f"""Printable {keyword} planner - print and start organizing today!

WHAT YOU GET:
- 1x Printable PDF Planner (A4 size)
- Cover page with table of contents
- {len(concept.get('pdf', {}).get('sections', []))} organized sections
- Notes pages for extra planning
- Clean, professional design

HOW TO USE:
1. Download the PDF after purchase
2. Print at home or at a print shop (A4 size recommended)
3. Fill in your plans and track your progress
4. Pair with our Excel template for digital + analog planning!

INSTANT DOWNLOAD - No physical item will be shipped.

© {datetime.now().year} {brand}"""

        writer.writerow({
            "title": f"{base_title} Printable Planner PDF Digital Download",
            "description": pdf_desc,
            "price": f"{pricing.get('pdf', 4.99):.2f}",
            "currency": "USD",
            "tags": tags_str,
            "materials": "",
            "quantity": "999",
            "sku": f"TM-PDF-{re.sub(r'[^A-Z0-9]', '', keyword.upper()[:20])}",
            "file": Path(files["pdf"]).name if files.get("pdf") else "",
            "preview_image": Path(files["preview"]).name if files.get("preview") else "",
            "category": "Craft Supplies & Tools > Templates",
            "type": "digital",
            "who_made": "i_did",
            "when_made": "2020_2025",
            "is_supply": "true",
        })

    # Bundle listing
    if files.get("excel") and files.get("pdf"):
        bundle_discount = int(config["pricing"]["bundle_discount"] * 100)
        bundle_desc = f"""BUNDLE & SAVE {bundle_discount}%! Get both the Excel template AND printable PDF planner.

{description}

WHAT YOU GET:
- 1x Excel Template (.xlsx) with formulas and dashboard
- 1x Printable PDF Planner (A4)
- Professional design, ready to use immediately

WHY BUNDLE?
- Save {bundle_discount}% compared to buying separately
- Digital + physical planning combo
- Track on your computer AND on paper

INSTANT DOWNLOAD - No physical item will be shipped.

© {datetime.now().year} {brand}"""

        writer.writerow({
            "title": f"{base_title} Bundle Excel + PDF Planner Digital Download",
            "description": bundle_desc,
            "price": f"{pricing.get('bundle', 13.99):.2f}",
            "currency": "USD",
            "tags": tags_str,
            "materials": "",
            "quantity": "999",
            "sku": f"TM-BUNDLE-{re.sub(r'[^A-Z0-9]', '', keyword.upper()[:20])}",
            "file": "",
            "preview_image": Path(files["preview"]).name if files.get("preview") else "",
            "category": "Craft Supplies & Tools > Templates",
            "type": "digital",
            "who_made": "i_did",
            "when_made": "2020_2025",
            "is_supply": "true",
        })


CSV_FIELDS = [
    "title", "description", "price", "currency", "tags", "materials",
    "quantity", "sku", "file", "preview_image", "category", "type",
    "who_made", "when_made", "is_supply",
]


def export_to_csv(products, config, output_path=None):
    """Ürün listesini Etsy bulk upload CSV'sine dönüştürür."""
    if output_path is None:
        output_dir = Path(config["output_dir"])
        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = output_dir / f"etsy_listings_{timestamp}.csv"

    with open(str(output_path), "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()

        if isinstance(products, dict):
            products = [products]

        for product in products:
            export_single(product, config, writer)

    count = len(products)
    logger.info(f"📋 Etsy CSV oluşturuldu: {output_path} ({count * 3} listing)")
    return str(output_path)
