"""Charge les JSON dans Supabase via HTTPS (port 443)."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import tools._bootstrap  # noqa: F401

from app.config import settings
from app import supabase_rest
from app.snapshot_codec import compress_json


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-id", default="default")
    parser.add_argument("--data-dir", type=Path, default=settings.local_data_path)
    args = parser.parse_args()

    if not settings.use_supabase_rest:
        raise SystemExit(
            "Remplissez SUPABASE_URL et SUPABASE_SERVICE_KEY dans api/.env\n"
            "  URL = https://VOTRE_REF.supabase.co\n"
            "  KEY = service_role (Settings - API Keys)"
        )

    files = {
        "BD": args.data_dir / "bd-sheet.json",
        "SYNTHESIS": args.data_dir / "synthesis-sheet.json",
    }
    for sheet, path in files.items():
        if not path.is_file():
            raise SystemExit(f"Missing {path}")
        print(f"Ingest {path.name} … (peut prendre 5-15 min, ne pas fermer)")
        with path.open(encoding="utf-8") as f:
            raw = json.load(f)
        total = len(raw.get("cells") or [])
        print(f"  {total} cellules a envoyer…")

        def on_batch(done: int) -> None:
            pct = min(100, round((done / total) * 100)) if total else 0
            print(f"  … {done}/{total} ({pct}%)", flush=True)

        n = supabase_rest.ingest_sheet(
            args.project_id, sheet, raw, progress=on_batch
        )
        print(f"  {sheet}: {n} cells OK")

    bd_path = files["BD"]
    syn_path = files["SYNTHESIS"]
    with bd_path.open(encoding="utf-8") as f:
        bd_raw = json.load(f)
    with syn_path.open(encoding="utf-8") as f:
        syn_raw = json.load(f)
    print("Session snapshot …")
    supabase_rest.upsert_session(
        args.project_id, 0, bd_raw, syn_raw, "ingest", structure_revision=0
    )
    print("Done — Supabase pret.")


if __name__ == "__main__":
    main()
