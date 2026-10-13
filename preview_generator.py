"""
PREVIEW GENERATOR — Etsy/Gumroad listing görseli (2000x1600)
=============================================================
Gerçek Excel dosyasını LibreOffice ile PDF'e, pdftoppm ile PNG'ye çevirip
Dashboard + ilk veri sayfasını satış görseline yerleştirir.
LibreOffice/pdftoppm yoksa çizilmiş mock tabloya düşer.
"""
import logging, re, shutil, subprocess, tempfile
from pathlib import Path

logger = logging.getLogger("template-machine.preview")

THEME_COLORS = {
    "ocean":    {"bg": (11, 37, 69),  "accent": (27, 73, 101),  "highlight": (95, 168, 211), "text": (255, 255, 255), "muted": (190, 233, 232)},
    "sunset":   {"bg": (45, 27, 61),  "accent": (108, 46, 122), "highlight": (212, 99, 138), "text": (255, 255, 255), "muted": (255, 236, 210)},
    "forest":   {"bg": (13, 40, 24),  "accent": (26, 71, 42),   "highlight": (74, 154, 91),  "text": (255, 255, 255), "muted": (212, 237, 218)},
    "midnight": {"bg": (10, 10, 26),  "accent": (26, 26, 46),   "highlight": (233, 69, 96),  "text": (255, 255, 255), "muted": (255, 224, 230)},
}
W, H = 2000, 1600

FONT_CANDIDATES = {
    "bold": ["arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
             "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"],
    "regular": ["arial.ttf", "Arial.ttf", "DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"],
}


def _font(size, bold=False):
    from PIL import ImageFont
    for name in FONT_CANDIDATES["bold" if bold else "regular"]:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def render_workbook_pages(xlsx_path, dpi=110, max_pages=4):
    """xlsx -> [PNG path, ...] (sayfa başına bir görsel). Araç yoksa []."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    pdftoppm = shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        logger.warning("LibreOffice/pdftoppm bulunamadı, önizleme mock tabloyla üretilecek")
        return []
    tmp = Path(tempfile.mkdtemp(prefix="tm_preview_"))
    try:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(xlsx_path)],
                       check=True, capture_output=True, timeout=180)
        pdf = tmp / (Path(xlsx_path).stem + ".pdf")
        subprocess.run([pdftoppm, "-png", "-r", str(dpi), "-l", str(max_pages), str(pdf), str(tmp / "page")],
                       check=True, capture_output=True, timeout=120)
        return sorted(str(p) for p in tmp.glob("page*.png"))
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError) as e:
        logger.warning(f"Excel render edilemedi: {e}")
        return []


def _autocrop(img, pad=18):
    from PIL import ImageChops, Image
    bg = Image.new(img.mode, img.size, (255, 255, 255))
    bbox = ImageChops.difference(img, bg).getbbox()
    if not bbox:
        return img
    l, t, r, b = bbox
    l, t = max(0, l - pad), max(0, t - pad)
    r, b = min(img.width, r + pad), min(img.height, b + pad)
    return img.crop((l, t, r, b))


def _paste_screenshot(canvas, shot, x, y, width, radius=14):
    """Beyaz çerçeve + gölge ile ekran görüntüsünü yerleştirir."""
    from PIL import Image, ImageDraw, ImageFilter
    ratio = width / shot.width
    shot = shot.resize((width, int(shot.height * ratio)), Image.LANCZOS)
    pad = 10
    frame_w, frame_h = shot.width + 2 * pad, shot.height + 2 * pad
    shadow = Image.new("RGBA", (frame_w + 80, frame_h + 80), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle([40, 48, 40 + frame_w, 48 + frame_h], radius=radius, fill=(0, 0, 0, 120))
    shadow = shadow.filter(ImageFilter.GaussianBlur(18))
    canvas.alpha_composite(shadow, (x - 40, y - 40))
    frame = Image.new("RGBA", (frame_w, frame_h), (255, 255, 255, 255))
    mask = Image.new("L", (frame_w, frame_h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, frame_w - 1, frame_h - 1], radius=radius, fill=255)
    frame.paste(shot.convert("RGBA"), (pad, pad))
    canvas.paste(frame, (x, y), mask)
    return y + frame_h


def _wrap(draw, text, font, max_w):
    words, lines, cur = text.split(), [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if draw.textlength(trial, font=font) > max_w and cur:
            lines.append(cur)
            cur = w
        else:
            cur = trial
    if cur:
        lines.append(cur)
    return lines


def generate_preview(keyword, spec, theme_name="ocean", output_dir="./products", excel_path=None):
    """Listing görselini üretir ve yolunu döndürür."""
    from PIL import Image, ImageDraw
    tc = THEME_COLORS.get(theme_name, THEME_COLORS["ocean"])
    canvas = Image.new("RGBA", (W, H), tc["bg"] + (255,))
    draw = ImageDraw.Draw(canvas)

    # Üst ve alt şerit
    draw.rectangle([(0, 0), (W, 90)], fill=tc["accent"])
    draw.rectangle([(0, H - 90), (W, H)], fill=tc["accent"])
    category = (spec.get("category") or "template").replace("_", " ").upper()
    draw.text((70, 30), f"{category} TEMPLATE", fill=tc["muted"], font=_font(26))
    draw.text((70, H - 62), "EXCEL  •  GOOGLE SHEETS  •  PRINTABLE PDF", fill=tc["highlight"], font=_font(28, bold=True))
    draw.text((W - 430, H - 60), "INSTANT DOWNLOAD", fill=tc["muted"], font=_font(26))

    # Sol metin bloğu
    left_w = 760
    y = 190
    for line in _wrap(draw, spec.get("product_name", keyword.title()), _font(68, bold=True), left_w)[:3]:
        draw.text((70, y), line, fill=tc["text"], font=_font(68, bold=True))
        y += 80
    y += 10
    draw.rectangle([(70, y), (330, y + 6)], fill=tc["highlight"])
    y += 34
    for line in _wrap(draw, spec.get("tagline", ""), _font(30), left_w)[:3]:
        draw.text((70, y), line, fill=tc["muted"], font=_font(30))
        y += 40
    y += 36
    bullets = [f"{sh['name']} tab" for sh in spec["excel"]["sheets"][:4]]
    bullets += ["Dashboard with live totals & charts", f"{len(spec['pdf']['pages'])}-page printable PDF planner"]
    for b in bullets:
        draw.ellipse([(78, y + 10), (96, y + 28)], fill=tc["highlight"])
        draw.text((116, y), b, fill=tc["text"], font=_font(28))
        y += 48

    # Sağ: gerçek ekran görüntüleri (Dashboard + ilk veri sayfası), yoksa mock
    pages = render_workbook_pages(excel_path) if excel_path else []
    if len(pages) >= 3:
        dash = _autocrop(Image.open(pages[1]).convert("RGB"))
        data = _autocrop(Image.open(pages[2]).convert("RGB"))
        bottom = _paste_screenshot(canvas, dash, 840, 150, 1100)
        _paste_screenshot(canvas, data, 960, min(bottom + 40, 880), 1000)
        draw = ImageDraw.Draw(canvas)
    else:
        _draw_mock_table(draw, spec, tc)

    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    safe_name = re.sub(r'[^\w\s-]', '', keyword).strip().replace(' ', '_')
    filepath = out_path / f"{safe_name}_preview.png"
    canvas.convert("RGB").save(str(filepath), "PNG", optimize=True)
    logger.info(f"🖼️ Önizleme oluşturuldu: {filepath} ({'gerçek ekran görüntüsü' if len(pages) >= 3 else 'mock tablo'})")
    return str(filepath)


def _draw_mock_table(draw, spec, tc):
    mx, my, mw, mh = 900, 200, 980, 1100
    draw.rectangle([(mx + 8, my + 8), (mx + mw + 8, my + mh + 8)], fill=(0, 0, 0))
    draw.rectangle([(mx, my), (mx + mw, my + mh)], fill=(255, 255, 255))
    draw.rectangle([(mx, my), (mx + mw, my + 60)], fill=tc["accent"])
    cols = [c["name"] for c in spec["excel"]["sheets"][0]["columns"][:5]]
    col_w = mw // max(len(cols), 1)
    for i, col in enumerate(cols):
        draw.text((mx + i * col_w + 16, my + 16), col[:14], fill=(255, 255, 255), font=_font(24, bold=True))
    rows = spec["excel"]["sheets"][0].get("rows", [])
    for r in range(16):
        ry = my + 60 + r * 64
        draw.rectangle([(mx, ry), (mx + mw, ry + 64)], fill=(245, 247, 250) if r % 2 == 0 else (255, 255, 255))
        if r < len(rows):
            for i, val in enumerate(rows[r][:len(cols)]):
                draw.text((mx + i * col_w + 16, ry + 18), str(val)[:16], fill=(40, 40, 60), font=_font(22))
    for i in range(1, len(cols)):
        draw.line([(mx + i * col_w, my + 60), (mx + i * col_w, my + mh)], fill=(220, 225, 230), width=2)
