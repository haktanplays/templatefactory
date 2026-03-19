"""
TEMPLATE MACHINE — Main Orchestrator
=====================================
Full pipeline: Discover → Generate → Publish → Report

Usage:
  python main.py                    # Full pipeline
  python main.py --trends-only      # Only discover trends
  python main.py --keyword "meal prep planner"  # Generate for specific keyword
  python main.py --status           # Show stats
"""
import json, sys, logging, argparse
from datetime import datetime
from pathlib import Path

# Setup logging
LOG_DIR = Path(__file__).parent
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_DIR / "template_machine.log", encoding="utf-8"),
    ]
)
logger = logging.getLogger("template-machine")

def load_config():
    from config_loader import load_config as _load
    return _load()

def save_run_log(results):
    log_path = Path(__file__).parent / "run_history.json"
    history = []
    if log_path.exists():
        try:
            with open(log_path) as f:
                history = json.load(f)
        except:
            history = []
    history.append({
        "timestamp": datetime.now().isoformat(),
        "results": results,
    })
    # Son 100 çalışmayı tut
    history = history[-100:]
    with open(log_path, "w") as f:
        json.dump(history, f, indent=2, ensure_ascii=False)

def get_past_keywords():
    """Daha önce üretilmiş keyword'leri döndürür."""
    log_path = Path(__file__).parent / "run_history.json"
    if not log_path.exists():
        return set()
    try:
        with open(log_path) as f:
            history = json.load(f)
        keywords = set()
        for run in history:
            for r in run.get("results", []):
                kw = r.get("keyword", "").lower().strip()
                if kw:
                    keywords.add(kw)
        return keywords
    except:
        return set()

def filter_duplicates(opportunities):
    """Daha önce üretilen keyword'leri fırsatlardan çıkarır."""
    past = get_past_keywords()
    if not past:
        return opportunities
    filtered = [o for o in opportunities if o["keyword"].lower().strip() not in past]
    skipped = len(opportunities) - len(filtered)
    if skipped:
        logger.info(f"  ♻️ {skipped} keyword daha önce üretilmiş, atlandı")
    return filtered

# ─── PIPELINE ───────────────────────────────────────────
def run_full_pipeline():
    """Tam pipeline: Discover → Generate → Publish → Report"""
    config = load_config()
    start_time = datetime.now()

    logger.info("=" * 60)
    logger.info("🦞 TEMPLATE MACHINE — Pipeline başlatılıyor")
    logger.info("=" * 60)

    # STEP 1: Trend Discovery
    logger.info("\n📊 ADIM 1: Trend keşfi...")
    from trend_scraper import find_opportunities
    opportunities = find_opportunities()

    if not opportunities:
        logger.warning("⚠️ Fırsat bulunamadı. Pipeline durduruluyor.")
        return {"status": "no_opportunities", "products": []}

    # Daha önce üretilenleri filtrele
    opportunities = filter_duplicates(opportunities)
    if not opportunities:
        logger.warning("⚠️ Tüm fırsatlar daha önce üretilmiş. Pipeline durduruluyor.")
        return {"status": "all_duplicates", "products": []}

    logger.info(f"\n🏆 {len(opportunities)} fırsat bulundu:")
    for i, opp in enumerate(opportunities):
        logger.info(f"  {i+1}. [{opp['score']}] {opp['keyword']} — ${opp['suggested_price_excel']}")

    # STEP 2: Product Generation
    logger.info("\n🏭 ADIM 2: Ürün üretimi...")
    from product_generator import generate_products
    all_products = []
    for opp in opportunities:
        try:
            product = generate_products(opp["keyword"], config)
            product["opportunity"] = opp
            all_products.append(product)
            logger.info(f"  ✅ {opp['keyword']}: Excel + PDF üretildi")
        except Exception as e:
            logger.error(f"  ❌ {opp['keyword']} üretim hatası: {e}")

    if not all_products:
        logger.error("Hiçbir ürün üretilemedi!")
        return {"status": "generation_failed", "products": []}

    # STEP 3: Gumroad Publishing
    logger.info("\n🚀 ADIM 3: Gumroad'a yükleme...")
    from gumroad_publisher import publish_products
    publish_results = []
    for product in all_products:
        try:
            result = publish_products(product, config)
            publish_results.append({
                "keyword": product["keyword"],
                "publish_status": result["status"],
                "files": product["files"],
                "pricing": product["pricing"],
                "listing_title": product.get("listing", {}).get("title", ""),
            })
        except Exception as e:
            logger.error(f"  ❌ {product['keyword']} yükleme hatası: {e}")
            publish_results.append({
                "keyword": product["keyword"],
                "publish_status": "error",
                "error": str(e),
            })

    # STEP 4: Etsy CSV Export
    logger.info("\n📋 ADIM 4: Etsy CSV export...")
    try:
        from etsy_exporter import export_to_csv
        csv_path = export_to_csv(all_products, config)
        logger.info(f"  ✅ Etsy CSV: {csv_path}")
    except Exception as e:
        logger.error(f"  ❌ Etsy CSV hatası: {e}")

    # STEP 5: Report
    elapsed = (datetime.now() - start_time).total_seconds()
    report = generate_report(publish_results, elapsed)
    save_run_log(publish_results)

    logger.info("\n" + report)
    return {"status": "complete", "products": publish_results, "report": report}

def run_trends_only():
    """Sadece trend keşfi yap, ürün üretme."""
    logger.info("🔍 Sadece trend keşfi modu...")
    from trend_scraper import find_opportunities
    opportunities = find_opportunities()

    logger.info(f"\n{'='*50}")
    logger.info("📊 TREND RAPORU")
    logger.info(f"{'='*50}")
    for i, opp in enumerate(opportunities):
        direction = "🔼" if opp["trend_direction"] == "rising" else "🔽" if opp["trend_direction"] == "falling" else "➡️"
        logger.info(f"\n  {i+1}. {opp['keyword']}")
        logger.info(f"     Skor: {opp['score']} | Trend: {direction} | Excel: ${opp['suggested_price_excel']} | PDF: ${opp['suggested_price_pdf']}")

    return opportunities

def run_single_keyword(keyword):
    """Tek bir keyword için ürün üret."""
    config = load_config()
    logger.info(f"🎯 Tek keyword modu: '{keyword}'")

    from product_generator import generate_products
    product = generate_products(keyword, config)

    from gumroad_publisher import publish_products
    result = publish_products(product, config)

    # Etsy CSV
    try:
        from etsy_exporter import export_to_csv
        csv_path = export_to_csv(product, config)
        logger.info(f"  📋 Etsy CSV: {csv_path}")
    except Exception as e:
        logger.error(f"  ❌ Etsy CSV hatası: {e}")

    logger.info(f"\n✅ Tamamlandı!")
    logger.info(f"  Excel: {product['files'].get('excel', 'N/A')}")
    logger.info(f"  PDF: {product['files'].get('pdf', 'N/A')}")
    logger.info(f"  Fiyat: Excel ${product['pricing']['excel']} | PDF ${product['pricing']['pdf']} | Bundle ${product['pricing']['bundle']}")
    logger.info(f"  Gumroad: {result['status']}")

    return product

def show_status():
    """Geçmiş çalışmaların istatistiklerini gösterir."""
    log_path = Path(__file__).parent / "run_history.json"
    if not log_path.exists():
        logger.info("📭 Henüz çalışma geçmişi yok.")
        return

    with open(log_path) as f:
        history = json.load(f)

    total_runs = len(history)
    total_products = sum(len(run.get("results", [])) for run in history)
    successful = sum(
        1 for run in history
        for r in run.get("results", [])
        if r.get("publish_status") == "published"
    )

    logger.info(f"\n{'='*50}")
    logger.info("📊 TEMPLATE MACHINE İSTATİSTİKLERİ")
    logger.info(f"{'='*50}")
    logger.info(f"  Toplam çalışma: {total_runs}")
    logger.info(f"  Toplam ürün: {total_products}")
    logger.info(f"  Başarılı yükleme: {successful}")
    logger.info(f"  Son çalışma: {history[-1]['timestamp'] if history else 'N/A'}")

# ─── REPORT ─────────────────────────────────────────────
def generate_report(results, elapsed_seconds):
    """Günlük özet raporu oluşturur."""
    success = [r for r in results if r.get("publish_status") in ["published", "skipped"]]
    failed = [r for r in results if r.get("publish_status") == "error"]

    total_excel_price = sum(r.get("pricing", {}).get("excel", 0) for r in success)
    total_pdf_price = sum(r.get("pricing", {}).get("pdf", 0) for r in success)
    total_bundle_price = sum(r.get("pricing", {}).get("bundle", 0) for r in success)

    lines = [
        "=" * 50,
        "🦞 TEMPLATE MACHINE — GÜNLÜK RAPOR",
        f"📅 {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "=" * 50,
        "",
        f"⏱️  Süre: {elapsed_seconds:.0f} saniye",
        f"✅ Başarılı: {len(success)} ürün seti",
        f"❌ Başarısız: {len(failed)}",
        "",
        "📦 ÜRÜNLER:",
    ]

    for r in success:
        kw = r.get("keyword", "?")
        ep = r.get("pricing", {}).get("excel", 0)
        pp = r.get("pricing", {}).get("pdf", 0)
        bp = r.get("pricing", {}).get("bundle", 0)
        lines.append(f"  • {kw}")
        lines.append(f"    Excel: ${ep} | PDF: ${pp} | Bundle: ${bp}")

    if failed:
        lines.append("\n⚠️ HATALAR:")
        for r in failed:
            lines.append(f"  • {r.get('keyword', '?')}: {r.get('error', 'unknown')}")

    lines.extend([
        "",
        f"💰 Potansiyel gelir (tümü satılırsa):",
        f"   Excel'ler: ${total_excel_price:.2f}",
        f"   PDF'ler: ${total_pdf_price:.2f}",
        f"   Bundle'lar: ${total_bundle_price:.2f}",
        "=" * 50,
    ])

    return "\n".join(lines)

# ─── ENTRY POINT ────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Template Machine — Automated Digital Product Factory")
    parser.add_argument("--trends-only", action="store_true", help="Only discover trends")
    parser.add_argument("--keyword", type=str, help="Generate for specific keyword")
    parser.add_argument("--status", action="store_true", help="Show statistics")
    args = parser.parse_args()

    if args.status:
        show_status()
    elif args.trends_only:
        run_trends_only()
    elif args.keyword:
        run_single_keyword(args.keyword)
    else:
        run_full_pipeline()
