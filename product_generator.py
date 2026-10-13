"""
PRODUCT GENERATOR — Keyword'den satılabilir Excel + PDF + listing görseli
=========================================================================
Akış: keyword -> kategori -> spec (Claude üretir; yoksa kategori paketi)
      -> excel_builder -> pdf_builder -> preview_generator
"""
import json, re, logging, time, copy
from pathlib import Path

from spec import detect_category, load_pack, validate_spec, SpecError, SCHEMA_DOC, supported_categories
from excel_builder import build_workbook
from pdf_builder import build_pdf
from preview_generator import generate_preview

logger = logging.getLogger("template-machine.generator")

SYSTEM_PROMPT = """You design premium digital planner products (Excel workbooks + printable PDF planners) that sell on Etsy and Gumroad.
Buyers pay for products that feel made for their exact situation: realistic sample data, the right columns, dropdowns, totals that
matter, and printable pages they would actually use. You output a single JSON product spec and nothing else."""


def load_config():
    from config_loader import load_config as _load
    return _load()


# ─── SPEC ÜRETİMİ ───────────────────────────────────────
def _extract_json(text):
    text = text.replace("```json", "").replace("```", "")
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        raise json.JSONDecodeError("JSON nesnesi bulunamadı", text, 0)
    return json.loads(text[start:end + 1])


def generate_spec(keyword, config):
    """(spec, category, source) döndürür. source: 'claude' | 'pack'."""
    category = detect_category(keyword)
    pack = load_pack(category) if category else None
    api_key = config.get("anthropic_api_key")

    if not api_key:
        if not pack:
            raise SpecError(f"'{keyword}' desteklenen kategorilerde değil ({supported_categories()}) ve ANTHROPIC_API_KEY yok")
        logger.warning(f"⚠️ ANTHROPIC_API_KEY yok: '{category}' kategori paketi olduğu gibi kullanılıyor")
        return validate_spec(copy.deepcopy(pack)), category, "pack"

    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    example = pack or load_pack("budget")
    user_prompt = f"""Design a complete product spec for this keyword.

KEYWORD: "{keyword}"
CATEGORY: {category or "general (closest example below)"}

{SCHEMA_DOC}

COMPLETE EXAMPLE SPEC (same format, related category — use it as a quality bar, do not copy it):
{json.dumps(example, ensure_ascii=False)}

Now write a NEW spec for "{keyword}". If the keyword emphasises one aspect (e.g. "guest list", "grocery", "debt payoff"),
make that aspect the first and richest sheet and the main dashboard KPIs. Pick a theme that fits the topic.
Output ONLY the JSON object."""

    last_error = None
    for attempt in range(3):
        try:
            msg = client.messages.create(
                model=config["model"],
                max_tokens=16000,
                output_config={"effort": "medium"},
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_prompt}],
            )
            if msg.stop_reason == "refusal":
                raise SpecError("Claude isteği reddetti")
            text = next(b.text for b in msg.content if b.type == "text")
            spec = validate_spec(_extract_json(text))
            spec.setdefault("category", category or "general")
            logger.info(f"✅ Claude spec üretti: {spec['product_name']} "
                        f"({msg.usage.input_tokens} in / {msg.usage.output_tokens} out token)")
            return spec, category, "claude"
        except (json.JSONDecodeError, SpecError, StopIteration) as e:
            last_error = e
            logger.warning(f"Claude spec geçersiz (deneme {attempt+1}/3): {e}")
        except anthropic.AuthenticationError:
            logger.error("Anthropic API anahtarı geçersiz")
            break
        except anthropic.BadRequestError as e:
            logger.error(f"Anthropic isteği reddedildi (model adı / parametre?): {e.message}")
            break
        except anthropic.RateLimitError as e:
            last_error = e
            wait = int(e.response.headers.get("retry-after", 15))
            logger.warning(f"Rate limit, {wait}s bekleniyor...")
            time.sleep(wait)
        except (anthropic.APIStatusError, anthropic.APIConnectionError) as e:
            last_error = e
            logger.warning(f"Anthropic API hatası (deneme {attempt+1}/3): {e}")
            time.sleep((attempt + 1) * 3)

    if pack:
        logger.error(f"Claude spec alınamadı ({last_error}); '{category}' kategori paketi kullanılıyor")
        return validate_spec(copy.deepcopy(pack)), category, "pack"
    raise SpecError(f"Spec üretilemedi: {last_error}")


# ─── DOSYA ÜRETİMİ ──────────────────────────────────────
def _safe_name(keyword):
    return re.sub(r'[^\w\s-]', '', keyword).strip().replace(' ', '_') or "template"


def generate_excel(keyword, spec, config, theme):
    output_dir = Path(config["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)
    brand = config["design"]["brand_name"]
    filepath = output_dir / f"{_safe_name(keyword)}_template.xlsx"
    build_workbook(spec, theme, brand=brand).save(str(filepath))
    # Önizleme için sayfa başına tek sayfa basan kopya (tablo boyu ne olursa olsun)
    preview_copy = output_dir / f".{_safe_name(keyword)}_preview_src.xlsx"
    build_workbook(spec, theme, brand=brand, fit_one_page=True).save(str(preview_copy))
    logger.info(f"📊 Excel oluşturuldu: {filepath} ({len(spec['excel']['sheets'])} veri sayfası, tema: {theme})")
    return str(filepath), str(preview_copy)


def generate_pdf(keyword, spec, config, theme):
    output_dir = Path(config["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)
    filepath = output_dir / f"{_safe_name(keyword)}_planner.pdf"
    return build_pdf(spec, theme, filepath, brand=config["design"]["brand_name"], tagline=config["design"]["tagline"])


def generate_products(keyword, config=None):
    """Bir keyword için tüm ürünleri üretir."""
    if config is None:
        config = load_config()
    logger.info(f"🏭 Ürün üretimi başlıyor: '{keyword}'")

    spec, category, source = generate_spec(keyword, config)
    themes = config["design"].get("themes", ["ocean"])
    theme = spec.get("theme") if spec.get("theme") in themes else themes[0]

    excel_path, preview_src = generate_excel(keyword, spec, config, theme)
    pdf_path = generate_pdf(keyword, spec, config, theme)
    preview_path = generate_preview(keyword, spec, theme_name=theme, output_dir=config["output_dir"], excel_path=preview_src)
    try:
        Path(preview_src).unlink()
    except OSError:
        pass

    excel_price = config["pricing"]["excel_base"]
    pdf_price = config["pricing"]["pdf_base"]
    bundle_price = round((excel_price + pdf_price) * (1 - config["pricing"]["bundle_discount"]), 2)

    return {
        "keyword": keyword,
        "category": category,
        "spec_source": source,
        "concept": spec,
        "theme": theme,
        "files": {"excel": excel_path, "pdf": pdf_path, "preview": preview_path},
        "pricing": {"excel": excel_price, "pdf": pdf_price, "bundle": bundle_price},
        "listing": spec.get("listing", {}),
    }


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    result = generate_products("meal prep planner")
    print(json.dumps({k: v for k, v in result.items() if k != "concept"}, indent=2))
