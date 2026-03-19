"""
PREVIEW GENERATOR — Creates product thumbnail images for Gumroad/Etsy listings
Uses: Pillow for image generation
"""
import logging
from pathlib import Path

logger = logging.getLogger("template-machine.preview")

# Tema renkleri (RGB)
THEME_COLORS = {
    "ocean":    {"bg": (11, 37, 69), "accent": (27, 73, 101), "highlight": (95, 168, 211), "text": (255, 255, 255), "muted": (190, 233, 232)},
    "sunset":   {"bg": (45, 27, 61), "accent": (108, 46, 122), "highlight": (212, 99, 138), "text": (255, 255, 255), "muted": (255, 236, 210)},
    "forest":   {"bg": (13, 40, 24), "accent": (26, 71, 42), "highlight": (74, 154, 91), "text": (255, 255, 255), "muted": (212, 237, 218)},
    "midnight": {"bg": (10, 10, 26), "accent": (26, 26, 46), "highlight": (233, 69, 96), "text": (255, 255, 255), "muted": (255, 224, 230)},
}


def generate_preview(keyword, concept, theme_name="ocean", output_dir="./products"):
    """Ürün için 1200x800 önizleme görseli oluşturur."""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        logger.warning("Pillow not installed, skipping preview generation")
        return None

    tc = THEME_COLORS.get(theme_name, THEME_COLORS["ocean"])
    W, H = 1200, 800

    img = Image.new("RGB", (W, H), tc["bg"])
    draw = ImageDraw.Draw(img)

    # Fontu yükle (sistem fontu yoksa default)
    def get_font(size, bold=False):
        font_names = ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf", "Helvetica.ttf"]
        if bold:
            font_names = ["arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"] + font_names
        for name in font_names:
            try:
                return ImageFont.truetype(name, size)
            except (OSError, IOError):
                continue
        return ImageFont.load_default()

    font_title = get_font(48, bold=True)
    font_subtitle = get_font(22)
    font_tag = get_font(16)
    font_brand = get_font(18, bold=True)

    # ── Üst banner ──
    draw.rectangle([(0, 0), (W, 120)], fill=tc["accent"])

    # Brand name
    brand_text = concept.get("category", "PREMIUM TEMPLATE").upper()
    draw.text((40, 45), brand_text, fill=tc["muted"], font=font_tag)

    # ── Product name ──
    product_name = concept.get("product_name", keyword.title())
    # Uzun isimleri sar
    if len(product_name) > 40:
        mid = len(product_name) // 2
        space_idx = product_name.rfind(" ", 0, mid + 10)
        if space_idx > 0:
            line1 = product_name[:space_idx]
            line2 = product_name[space_idx + 1:]
            draw.text((60, 160), line1, fill=tc["text"], font=font_title)
            draw.text((60, 220), line2, fill=tc["text"], font=font_title)
            y_after_title = 290
        else:
            draw.text((60, 170), product_name, fill=tc["text"], font=font_title)
            y_after_title = 240
    else:
        draw.text((60, 170), product_name, fill=tc["text"], font=font_title)
        y_after_title = 240

    # ── Accent line ──
    draw.rectangle([(60, y_after_title), (300, y_after_title + 4)], fill=tc["highlight"])

    # ── Tagline ──
    tagline = concept.get("tagline", f"Professional {keyword} template")
    draw.text((60, y_after_title + 20), tagline, fill=tc["muted"], font=font_subtitle)

    # ── Included items ──
    y = y_after_title + 80
    items = [
        "Excel Template with Formulas",
        "Printable PDF Planner",
        "Professional Dashboard",
        "Instant Digital Download",
    ]
    for item in items:
        # Bullet
        draw.ellipse([(70, y + 6), (82, y + 18)], fill=tc["highlight"])
        draw.text((95, y), item, fill=tc["text"], font=font_tag)
        y += 32

    # ── Sağ tarafta mock spreadsheet ──
    mock_x, mock_y = 680, 160
    mock_w, mock_h = 460, 500

    # Spreadsheet gölge
    draw.rectangle([(mock_x + 4, mock_y + 4), (mock_x + mock_w + 4, mock_y + mock_h + 4)],
                   fill=(0, 0, 0))
    # Spreadsheet arka planı
    draw.rectangle([(mock_x, mock_y), (mock_x + mock_w, mock_y + mock_h)],
                   fill=(255, 255, 255))

    # Header satırı
    draw.rectangle([(mock_x, mock_y), (mock_x + mock_w, mock_y + 35)],
                   fill=tc["accent"])
    cols = concept.get("excel", {}).get("sheets", [{}])[0].get("columns", ["A", "B", "C", "D"])[:4]
    col_w = mock_w // max(len(cols), 1)
    for i, col in enumerate(cols):
        text = col[:10]
        draw.text((mock_x + i * col_w + 10, mock_y + 8), text,
                  fill=(255, 255, 255), font=font_tag)

    # Data satırları (zebra)
    for row in range(12):
        ry = mock_y + 35 + row * 38
        bg = (245, 247, 250) if row % 2 == 0 else (255, 255, 255)
        draw.rectangle([(mock_x, ry), (mock_x + mock_w, ry + 38)], fill=bg)
        # Satır çizgisi
        draw.line([(mock_x, ry + 38), (mock_x + mock_w, ry + 38)], fill=(220, 225, 230), width=1)

    # Kolon çizgileri
    for i in range(1, len(cols)):
        cx = mock_x + i * col_w
        draw.line([(cx, mock_y + 35), (cx, mock_y + mock_h)], fill=(220, 225, 230), width=1)

    # ── Alt banner ──
    draw.rectangle([(0, H - 70), (W, H)], fill=tc["accent"])
    draw.text((40, H - 50), "EXCEL + PDF BUNDLE", fill=tc["highlight"], font=font_brand)
    draw.text((W - 250, H - 50), "INSTANT DOWNLOAD", fill=tc["muted"], font=font_tag)

    # Kaydet
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    import re
    safe_name = re.sub(r'[^\w\s-]', '', keyword).strip().replace(' ', '_')
    filename = f"{safe_name}_preview.png"
    filepath = out_path / filename
    img.save(str(filepath), "PNG", quality=95)
    logger.info(f"🖼️ Önizleme oluşturuldu: {filepath}")
    return str(filepath)
