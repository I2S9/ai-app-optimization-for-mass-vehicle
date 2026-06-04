"""Test connexion Supabase — lancer depuis api/ : .\\go-test.bat"""
from __future__ import annotations

import sys

import tools._bootstrap  # noqa: F401 — adds api/ to PYTHONPATH

from app.config import settings


def main() -> int:
    if not settings.use_postgres_store:
        print("DATA_BACKEND=postgres ou supabase + SUPABASE_DB_URL dans api/.env")
        return 1
    print("Mode:", settings.api_mode)
    print("URL host:", settings.pg_url.split("@")[-1].split("/")[0])
    try:
        from app.supabase_store import fetch_meta, pg_connection

        with pg_connection() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
                print("Connexion OK")
        meta = fetch_meta("default", "BD")
        if meta:
            print("Table sheet_meta OK — lastRow:", meta.get("lastRow"))
        else:
            print("Connexion OK mais pas de donnees — lancer ingest_json_to_supabase.py")
        return 0
    except Exception as e:
        err = str(e)
        print("ECHEC:", err)
        if "timed out" in err.lower() or "10060" in err:
            print()
            print("=" * 60)
            print("  RESEAU ENTREPRISE : Supabase est BLOQUE (port 6543).")
            print("  Ce n'est PAS votre faute. Mot de passe / SQL OK.")
            print()
            print("  DEMO MAINTENANT : double-clic run-bd-server.bat")
            print("  (matrix + calculs temps reel — sans cloud)")
            print()
            print("  Supabase plus tard : partage 4G / VPN, puis go-test a nouveau")
            print("=" * 60)
        else:
            print("\nAide:")
            print("  1. Supabase - Connect - copier URI Session pooler")
            print("  2. Mot de passe avec & * = : encoder en URL ou reset simple")
        return 1


if __name__ == "__main__":
    sys.exit(main())
