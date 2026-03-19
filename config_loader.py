"""
CONFIG LOADER — Merkezi konfigürasyon yükleyici
.env dosyasından API key'leri, config.json'dan ayarları okur.
"""
import json, os
from pathlib import Path

def load_config():
    """Config.json + .env'den birleşik konfigürasyon yükler."""
    # .env dosyasını yükle (varsa)
    env_path = Path(__file__).parent / ".env"
    if env_path.exists():
        try:
            from dotenv import load_dotenv
            load_dotenv(env_path)
        except ImportError:
            # python-dotenv yoksa manuel oku
            with open(env_path) as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        key, _, value = line.partition("=")
                        os.environ[key.strip()] = value.strip()

    # config.json'u oku
    cfg_path = Path(__file__).parent / "config.json"
    with open(cfg_path) as f:
        config = json.load(f)

    # API key'leri .env / ortam değişkenlerinden al
    config["anthropic_api_key"] = os.environ.get("ANTHROPIC_API_KEY", "")
    config["gumroad_access_token"] = os.environ.get("GUMROAD_ACCESS_TOKEN", "")

    return config
