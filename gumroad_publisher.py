"""
GUMROAD PUBLISHER — Uploads products to Gumroad via API
Docs: https://app.gumroad.com/api
"""
import json, logging, zipfile, tempfile, requests
from pathlib import Path

logger = logging.getLogger("template-machine.gumroad")

def load_config():
    from config_loader import load_config as _load
    return _load()

GUMROAD_API = "https://api.gumroad.com/v2"

def create_product(name, description, price_cents, tags, config):
    """Gumroad'da yeni ürün oluşturur."""
    token = config["gumroad_access_token"]
    data = {
        "access_token": token,
        "name": name,
        "description": description,
        "price": price_cents,
        "currency": "usd",
        "tags": ",".join(tags[:5]) if tags else "",
        "published": not config["gumroad"].get("draft_first", True),
    }
    try:
        r = requests.post(f"{GUMROAD_API}/products", data=data, timeout=30)
        if r.status_code == 200:
            product = r.json().get("product", {})
            logger.info(f"✅ Gumroad ürün oluşturuldu: {product.get('name')} (ID: {product.get('id')})")
            return product
        else:
            logger.error(f"❌ Gumroad ürün oluşturma hatası: {r.status_code} — {r.text}")
            return None
    except Exception as e:
        logger.error(f"❌ Gumroad API hatası: {e}")
        return None

def upload_file(product_id, filepath, config):
    """Ürüne dosya ekler."""
    token = config["gumroad_access_token"]
    try:
        with open(filepath, "rb") as f:
            files = {"file": (Path(filepath).name, f)}
            data = {"access_token": token}
            r = requests.put(f"{GUMROAD_API}/products/{product_id}", data=data, files=files, timeout=60)
            if r.status_code == 200:
                logger.info(f"📎 Dosya yüklendi: {Path(filepath).name}")
                return True
            else:
                logger.error(f"❌ Dosya yükleme hatası: {r.status_code}")
                return False
    except Exception as e:
        logger.error(f"❌ Dosya yükleme hatası: {e}")
        return False

def upload_thumbnail(product_id, image_path, config):
    """Ürüne önizleme görseli (thumbnail) ekler."""
    if not image_path:
        return False
    token = config["gumroad_access_token"]
    try:
        with open(image_path, "rb") as f:
            files = {"preview": (Path(image_path).name, f, "image/png")}
            data = {"access_token": token}
            r = requests.put(f"{GUMROAD_API}/products/{product_id}", data=data, files=files, timeout=60)
            if r.status_code == 200:
                logger.info(f"🖼️ Thumbnail yüklendi: {Path(image_path).name}")
                return True
            else:
                logger.error(f"❌ Thumbnail yükleme hatası: {r.status_code}")
                return False
    except Exception as e:
        logger.error(f"❌ Thumbnail yükleme hatası: {e}")
        return False

def create_bundle_zip(excel_path, pdf_path):
    """Bundle için Excel + PDF'i tek bir ZIP dosyasına paketler."""
    zip_path = Path(tempfile.mkdtemp()) / f"{Path(excel_path).stem}_bundle.zip"
    with zipfile.ZipFile(str(zip_path), "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(excel_path, Path(excel_path).name)
        zf.write(pdf_path, Path(pdf_path).name)
    logger.info(f"📦 Bundle ZIP oluşturuldu: {zip_path}")
    return str(zip_path)

def publish_product(product_id, config):
    """Draft ürünü yayına alır."""
    token = config["gumroad_access_token"]
    try:
        r = requests.put(f"{GUMROAD_API}/products/{product_id}",
                        data={"access_token": token, "published": True}, timeout=30)
        if r.status_code == 200:
            logger.info(f"🚀 Ürün yayında!")
            return True
        return False
    except Exception as e:
        logger.error(f"❌ Yayınlama hatası: {e}")
        return False

def _publish_single(product_id, config):
    """draft_first False ise ürünü yayınlar."""
    if not config["gumroad"].get("draft_first", True):
        publish_product(product_id, config)

def publish_products(product_data, config=None):
    """Üretilen ürünleri Gumroad'a yükler."""
    if config is None:
        config = load_config()

    if not config.get("gumroad_access_token"):
        logger.warning("⚠️ Gumroad token ayarlanmamış! Ürünler sadece lokal kaydedildi.")
        return {"status": "skipped", "reason": "no_token", "files": product_data.get("files", {})}

    results = []
    listing = product_data.get("listing", {})
    files = product_data.get("files", {})
    pricing = product_data.get("pricing", {})
    preview = files.get("preview")

    # Excel ürünü
    if files.get("excel"):
        name = f"{listing.get('title', product_data['keyword'].title() + ' Template')} — Excel"
        desc = listing.get("description", f"Professional {product_data['keyword']} Excel template")
        tags = listing.get("tags", [])
        price = int(pricing.get("excel", 14.99) * 100)

        product = create_product(name, desc, price, tags, config)
        if product:
            upload_file(product["id"], files["excel"], config)
            upload_thumbnail(product["id"], preview, config)
            _publish_single(product["id"], config)
            results.append({"type": "excel", "id": product["id"], "url": product.get("short_url", ""), "price": pricing["excel"]})

    # PDF ürünü
    if files.get("pdf"):
        name = f"{listing.get('title', product_data['keyword'].title() + ' Planner')} — Printable PDF"
        desc = f"Printable {product_data['keyword']} planner. Print and use daily. Perfect companion to the Excel template."
        tags = listing.get("tags", [])
        price = int(pricing.get("pdf", 4.99) * 100)

        product = create_product(name, desc, price, tags, config)
        if product:
            upload_file(product["id"], files["pdf"], config)
            upload_thumbnail(product["id"], preview, config)
            _publish_single(product["id"], config)
            results.append({"type": "pdf", "id": product["id"], "url": product.get("short_url", ""), "price": pricing["pdf"]})

    # Bundle — ZIP ile tek dosya olarak yükle
    if files.get("excel") and files.get("pdf"):
        name = f"{listing.get('title', product_data['keyword'].title())} — Complete Bundle (Excel + PDF)"
        desc = f"BUNDLE DEAL! Get both the Excel template AND printable PDF planner. Save {int(config['pricing']['bundle_discount']*100)}%!"
        tags = listing.get("tags", []) + ["bundle", "deal"]
        price = int(pricing.get("bundle", 13.99) * 100)

        product = create_product(name, desc, price, tags[:5], config)
        if product:
            bundle_zip = create_bundle_zip(files["excel"], files["pdf"])
            upload_file(product["id"], bundle_zip, config)
            upload_thumbnail(product["id"], preview, config)
            _publish_single(product["id"], config)
            results.append({"type": "bundle", "id": product["id"], "url": product.get("short_url", ""), "price": pricing["bundle"]})
            # Temp ZIP temizliği
            try:
                Path(bundle_zip).unlink()
            except OSError:
                pass

    logger.info(f"📦 {len(results)} ürün Gumroad'a yüklendi")
    return {"status": "published", "products": results}

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    print("Gumroad publisher hazır. publish_products() ile kullanın.")
