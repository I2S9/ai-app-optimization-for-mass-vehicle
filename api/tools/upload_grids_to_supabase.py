"""Upload precomputed display grid packs (*-grid.json) into Supabase via HTTPS.

Build the packs first:  node tools/build-grid-sheets.mjs
Then upload them:        python -m tools.upload_grids_to_supabase

This is what lets the app serve the precomputed grid from the DB
(GET /api/v1/sheets/{id}/grid) instead of the static /public/data/*-grid.json file.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import tools._bootstrap  # noqa: F401

from app.config import settings
from app import supabase_rest


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-id", default="default")
    parser.add_argument("--data-dir", type=Path, default=settings.local_data_path)
    args = parser.parse_args()

    if not settings.use_supabase_rest:
        raise SystemExit(
            "Remplissez SUPABASE_URL et SUPABASE_SERVICE_KEY dans api/.env"
        )

    files = {
        "BD": args.data_dir / "bd-sheet-grid.json",
        "SYNTHESIS": args.data_dir / "synthesis-sheet-grid.json",
    }
    for sheet, path in files.items():
        if not path.is_file():
            raise SystemExit(
                f"Missing {path} — lancez d'abord: node tools/build-grid-sheets.mjs"
            )
        size_mb = path.stat().st_size / (1024 * 1024)
        print(f"Upload {path.name} ({size_mb:.1f} MB) -> sheet_grid_cache[{sheet}] …")
        with path.open(encoding="utf-8") as f:
            pack = json.load(f)
        fingerprint = str(pack.get("fingerprint") or "")
        supabase_rest.upsert_grid_pack(args.project_id, sheet, fingerprint, pack)
        print(f"  {sheet}: OK (fingerprint {fingerprint})")

    print("Done — grilles servies depuis Supabase.")


if __name__ == "__main__":
    main()
