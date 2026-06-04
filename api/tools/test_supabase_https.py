"""Test Supabase HTTPS — python tools/test_supabase_https.py"""
from __future__ import annotations

import sys

import tools._bootstrap  # noqa: F401

from app.config import settings
from app import supabase_rest


def main() -> int:
    if not settings.use_supabase_rest:
        print("Remplissez SUPABASE_URL + SUPABASE_SERVICE_KEY dans api/.env")
        return 1
    print("URL:", settings.supabase_url)
    try:
        m = supabase_rest.fetch_meta("default", "BD")
    except Exception as e:
        print("ECHEC connexion Supabase:", e)
        print()
        print("Ne pas ouvrir https://xxx.supabase.co dans le navigateur pour l'app.")
        print("URLs correctes:")
        print("  - Application : http://127.0.0.1:5173/")
        print("  - API config  : http://127.0.0.1:5173/api/v1/config")
        return 1
    if not m or m.get("lastRow") is None:
        print("Tables OK mais pas de donnees BD — lancez go-ingest-supabase.bat")
        return 1
    print("Connexion OK — meta BD lastRow =", m.get("lastRow"))
    cells = supabase_rest.fetch_cells("default", "BD", 2, 5)
    print("Echantillon cellules:", len(cells))
    return 0


if __name__ == "__main__":
    sys.exit(main())
