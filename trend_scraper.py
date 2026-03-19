"""
TREND SCRAPER — Discovers trending template opportunities
Uses: Google Trends (pytrends) + Etsy Autocomplete + Scoring Algorithm
"""
import json, time, re, logging
from datetime import datetime, timedelta
from pathlib import Path

logger = logging.getLogger("template-machine.trends")

def load_config():
    from config_loader import load_config as _load
    return _load()

# ─── ETSY AUTOCOMPLETE ─────────────────────────────────
def fetch_etsy_autocomplete(query, max_retries=2):
    """Etsy arama çubuğunun önerdiği tamamlamaları çeker."""
    import requests
    url = "https://www.etsy.com/api/v3/ajax/bespoke/member/neu/specs/autocomplete"
    params = {"query": query, "limit": 10}
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    for attempt in range(max_retries):
        try:
            r = requests.get(url, params=params, headers=headers, timeout=10)
            if r.status_code == 200:
                data = r.json()
                if isinstance(data, list):
                    return [item if isinstance(item, str) else item.get("query", "") for item in data]
                elif isinstance(data, dict) and "results" in data:
                    return [r.get("query", r) if isinstance(r, dict) else r for r in data["results"]]
            time.sleep(1)
        except Exception as e:
            logger.warning(f"Etsy autocomplete failed for '{query}': {e}")
            time.sleep(2)
    return []

def fetch_etsy_autocomplete_simple(query):
    """Fallback: basit GET ile Etsy önerilerini çeker."""
    import requests
    try:
        url = f"https://www.etsy.com/search/suggest?q={query}"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            try:
                return r.json() if isinstance(r.json(), list) else []
            except:
                return []
    except:
        pass
    return []

# ─── GOOGLE TRENDS ──────────────────────────────────────
def fetch_google_trends(keywords, timeframe="now 7-d"):
    """Google Trends'ten keyword ilgi skorlarını çeker."""
    try:
        from pytrends.request import TrendReq
        pytrends = TrendReq(hl="en-US", tz=360)
        results = {}
        # Google Trends max 5 keyword per request
        for i in range(0, len(keywords), 5):
            batch = keywords[i:i+5]
            try:
                pytrends.build_payload(batch, cat=0, timeframe=timeframe, geo="")
                interest = pytrends.interest_over_time()
                if not interest.empty:
                    for kw in batch:
                        if kw in interest.columns:
                            vals = interest[kw].tolist()
                            results[kw] = {
                                "current": vals[-1] if vals else 0,
                                "avg": sum(vals) / len(vals) if vals else 0,
                                "trend": "rising" if len(vals) > 1 and vals[-1] > vals[0] else "falling",
                                "peak": max(vals) if vals else 0,
                            }
                time.sleep(1)
            except Exception as e:
                logger.warning(f"Google Trends batch failed: {e}")
                for kw in batch:
                    results[kw] = {"current": 50, "avg": 50, "trend": "unknown", "peak": 50}
        return results
    except ImportError:
        logger.warning("pytrends not installed, using fallback scores")
        return {kw: {"current": 50, "avg": 50, "trend": "unknown", "peak": 50} for kw in keywords}

# ─── KEYWORD DISCOVERY ──────────────────────────────────
def discover_keywords(config):
    """Seed keyword'lerden trending template fırsatlarını keşfeder."""
    settings = config["trend_settings"]
    all_keywords = set()

    logger.info("🔍 Etsy autocomplete taranıyor...")
    for topic in settings["seed_topics"]:
        for pattern in settings["seed_keywords"]:
            query = pattern.replace("{keyword}", topic)
            results = fetch_etsy_autocomplete(query)
            if not results:
                results = fetch_etsy_autocomplete_simple(query)
            for r in results:
                if isinstance(r, str) and len(r) > 5:
                    cleaned = r.strip().lower()
                    if any(w in cleaned for w in ["template", "spreadsheet", "planner", "tracker", "checklist", "calculator", "dashboard", "budget"]):
                        all_keywords.add(cleaned)
            time.sleep(0.5)

    logger.info(f"📊 {len(all_keywords)} keyword bulundu, Google Trends ile skorlanıyor...")
    keywords_list = list(all_keywords)[:50]

    if keywords_list:
        trend_data = fetch_google_trends(keywords_list)
    else:
        trend_data = {}
        logger.warning("Etsy'den keyword bulunamadı, fallback kullanılıyor")
        for topic in settings["seed_topics"][:10]:
            for suffix in ["template", "planner", "tracker"]:
                kw = f"{topic} {suffix}"
                keywords_list.append(kw)
                trend_data[kw] = {"current": 50, "avg": 50, "trend": "unknown", "peak": 50}

    return keywords_list, trend_data

# ─── OPPORTUNITY SCORING ────────────────────────────────
def score_opportunities(keywords, trend_data, config):
    """Her keyword'e fırsat skoru atar."""
    scored = []
    for kw in keywords:
        td = trend_data.get(kw, {"current": 30, "avg": 30, "trend": "unknown", "peak": 30})

        # Trend skoru (0-100)
        trend_score = td["current"]
        if td["trend"] == "rising":
            trend_score *= 1.3
        elif td["trend"] == "falling":
            trend_score *= 0.7

        # Keyword kalitesi (uzunluk, spesifiklik)
        word_count = len(kw.split())
        specificity_bonus = min(word_count * 5, 25)

        # Template tipi bonusu
        type_bonus = 0
        high_value_types = ["calculator", "dashboard", "tracker", "budget", "planner"]
        for t in high_value_types:
            if t in kw:
                type_bonus = 15
                break

        # Final skor
        total_score = trend_score + specificity_bonus + type_bonus
        scored.append({
            "keyword": kw,
            "score": round(total_score, 1),
            "trend_score": round(trend_score, 1),
            "trend_direction": td["trend"],
            "specificity": word_count,
            "suggested_price_excel": round(config["pricing"]["excel_base"] * (1.2 if total_score > 80 else 1.0), 2),
            "suggested_price_pdf": round(config["pricing"]["pdf_base"] * (1.2 if total_score > 80 else 1.0), 2),
        })

    scored.sort(key=lambda x: x["score"], reverse=True)
    max_results = config["trend_settings"]["max_results"]
    top = scored[:max_results]
    logger.info(f"🏆 Top {len(top)} fırsat seçildi")
    for i, opp in enumerate(top):
        logger.info(f"  {i+1}. [{opp['score']}] {opp['keyword']} ({opp['trend_direction']})")
    return top

# ─── MAIN ───────────────────────────────────────────────
def find_opportunities():
    """Ana trend keşif fonksiyonu."""
    config = load_config()
    keywords, trend_data = discover_keywords(config)
    opportunities = score_opportunities(keywords, trend_data, config)
    return opportunities

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    opps = find_opportunities()
    print(json.dumps(opps, indent=2, ensure_ascii=False))
