#!/usr/bin/env python3
"""
TEMPLATE MACHINE — Kurulum Scripti
Çalıştır: python setup.py
"""
import subprocess, sys, os, json
from pathlib import Path

def run(cmd):
    print(f"  -> {cmd}")
    subprocess.run(cmd, shell=True, check=True)

def main():
    print("=" * 50)
    print("TEMPLATE MACHINE — KURULUM")
    print("=" * 50)

    # 1. Python paketleri (requirements.txt'ten)
    print("\n[1/4] Python paketleri yukleniyor...")
    req_path = Path(__file__).parent / "requirements.txt"
    if req_path.exists():
        run(f"{sys.executable} -m pip install -r {req_path}")
    else:
        packages = ["openpyxl", "reportlab", "pytrends", "requests", "anthropic", "python-dotenv", "Pillow"]
        for pkg in packages:
            try:
                __import__(pkg.replace("-", "_").split(">")[0])
                print(f"  [OK] {pkg} zaten yuklu")
            except ImportError:
                print(f"  [DL] {pkg} yukleniyor...")
                run(f"{sys.executable} -m pip install {pkg}")

    # 2. Output dizini
    print("\n[2/4] Dizinler olusturuluyor...")
    output_dir = Path(__file__).parent / "products"
    output_dir.mkdir(exist_ok=True)
    print(f"  [OK] {output_dir}")

    # 3. .env dosyasi
    print("\n[3/4] API key ayarlari...")
    env_path = Path(__file__).parent / ".env"
    env_vars = {}

    if env_path.exists():
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, _, value = line.partition("=")
                    env_vars[key.strip()] = value.strip()
        print("  [OK] Mevcut .env dosyasi bulundu")

    if not env_vars.get("ANTHROPIC_API_KEY") or env_vars["ANTHROPIC_API_KEY"].startswith("sk-ant-your"):
        print("\n  Anthropic API key ayarlanmamis!")
        print("  -> https://console.anthropic.com/settings/keys adresinden key al")
        key = input("  Anthropic API Key: ").strip()
        if key:
            env_vars["ANTHROPIC_API_KEY"] = key

    if not env_vars.get("GUMROAD_ACCESS_TOKEN") or env_vars["GUMROAD_ACCESS_TOKEN"].startswith("your-"):
        print("\n  Gumroad access token ayarlanmamis!")
        print("  -> https://app.gumroad.com/settings/advanced#application-form adresinden token al")
        token = input("  Gumroad Access Token (bos birakirsan sonra girersin): ").strip()
        if token:
            env_vars["GUMROAD_ACCESS_TOKEN"] = token

    # .env dosyasini yaz
    with open(env_path, "w") as f:
        for k, v in env_vars.items():
            f.write(f"{k}={v}\n")
    print("  [OK] .env dosyasi guncellendi")

    # Marka adi
    config_path = Path(__file__).parent / "config.json"
    if config_path.exists():
        with open(config_path) as f:
            config = json.load(f)
        brand = input(f"\n  Marka adin ne olsun? (Enter = {config['design']['brand_name']}): ").strip()
        if brand:
            config["design"]["brand_name"] = brand
            with open(config_path, "w") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            print("  [OK] Config guncellendi")

    # 4. Test
    print("\n[4/4] Hizli test...")
    try:
        from config_loader import load_config
        cfg = load_config()
        has_key = bool(cfg.get("anthropic_api_key"))
        print(f"  [OK] Config okunabiliyor (API key: {'var' if has_key else 'YOK'})")
    except Exception as e:
        print(f"  [FAIL] Config hatasi: {e}")

    try:
        import openpyxl
        print("  [OK] openpyxl calisiyor")
    except ImportError:
        print("  [FAIL] openpyxl yuklenemedi")

    try:
        import reportlab
        print("  [OK] reportlab calisiyor")
    except ImportError:
        print("  [FAIL] reportlab yuklenemedi")

    # 5. Ozet
    print("\n" + "=" * 50)
    print("KURULUM TAMAMLANDI!")
    print("=" * 50)
    print("""
Kullanim:
  python main.py                          # Tam pipeline calistir
  python main.py --trends-only            # Sadece trendleri goster
  python main.py --keyword "meal prep"    # Tek keyword icin uret
  python main.py --status                 # Istatistikleri goster

Cron job (her gun 08:00):
  Linux/Mac: crontab -e -> 0 8 * * * cd /path/to/template-machine && python main.py
  Windows:   Task Scheduler ile ayarla
""")

if __name__ == "__main__":
    main()
