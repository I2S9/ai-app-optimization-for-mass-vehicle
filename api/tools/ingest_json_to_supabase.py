"""Ingest exported JSON into Supabase — lancer depuis api/ : .\\go-ingest.bat"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import tools._bootstrap  # noqa: F401

from psycopg2.extras import execute_batch

from app.config import settings
from app.snapshot_codec import compress_json
from app.supabase_store import ingest_connection


def ingest(project_id: str, data_dir: Path, *, snapshots_only: bool = False) -> None:
    files = {
        "BD": data_dir / "bd-sheet.json",
        "SYNTHESIS": data_dir / "synthesis-sheet.json",
    }
    with ingest_connection() as conn:
        with conn.cursor() as cur:
            for sheet, path in files.items():
                if not path.is_file():
                    print(f"Skip missing {path}")
                    continue
                print(f"Loading {path.name} …")
                with path.open(encoding="utf-8") as f:
                    raw = json.load(f)
                meta = {k: v for k, v in raw.items() if k != "cells"}
                cur.execute(
                    """
                    INSERT INTO public.sheet_meta (project_id, sheet, meta_json, updated_at)
                    VALUES (%s, %s, %s::jsonb, now())
                    ON CONFLICT (project_id, sheet) DO UPDATE SET
                      meta_json = EXCLUDED.meta_json,
                      updated_at = now()
                    """,
                    (project_id, sheet, json.dumps(meta, ensure_ascii=False)),
                )
                if snapshots_only:
                    print(f"  {sheet}: meta only ({len(raw.get('cells') or [])} cells)")
                    continue
                cur.execute(
                    "DELETE FROM public.sheet_cells WHERE project_id = %s AND sheet = %s",
                    (project_id, sheet),
                )
                batch = []
                for cell in raw.get("cells") or []:
                    batch.append(
                        (
                            project_id,
                            sheet,
                            int(cell["r"]),
                            str(cell["c"]),
                            str(cell.get("v", "")) if cell.get("v") is not None else None,
                            str(cell["f"]) if cell.get("f") else None,
                            bool(cell.get("userEdited")),
                        )
                    )
                    if len(batch) >= 5000:
                        execute_batch(
                            cur,
                            """
                            INSERT INTO public.sheet_cells
                              (project_id, sheet, excel_row, col, v, f, user_edited)
                            VALUES (%s, %s, %s, %s, %s, %s, %s)
                            ON CONFLICT (project_id, sheet, excel_row, col) DO UPDATE SET
                              v = EXCLUDED.v, f = EXCLUDED.f, user_edited = EXCLUDED.user_edited
                            """,
                            batch,
                        )
                        batch.clear()
                if batch:
                    execute_batch(
                        cur,
                        """
                        INSERT INTO public.sheet_cells
                          (project_id, sheet, excel_row, col, v, f, user_edited)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                        ON CONFLICT (project_id, sheet, excel_row, col) DO UPDATE SET
                          v = EXCLUDED.v, f = EXCLUDED.f, user_edited = EXCLUDED.user_edited
                        """,
                        batch,
                    )
                print(f"  {sheet}: {len(raw.get('cells') or [])} cells")

            bd_path = files["BD"]
            syn_path = files["SYNTHESIS"]
            if bd_path.is_file() and syn_path.is_file():
                with bd_path.open(encoding="utf-8") as f:
                    bd_raw = json.load(f)
                with syn_path.open(encoding="utf-8") as f:
                    syn_raw = json.load(f)
                print("Saving workbook_sessions (compressed) …")
                cur.execute(
                    """
                    INSERT INTO public.workbook_sessions (
                      project_id, revision, structure_revision,
                      bd_snapshot, syn_snapshot, updated_at, updated_by
                    )
                    VALUES (%s, 0, 0, %s, %s, now(), 'ingest')
                    ON CONFLICT (project_id) DO UPDATE SET
                      bd_snapshot = EXCLUDED.bd_snapshot,
                      syn_snapshot = EXCLUDED.syn_snapshot,
                      updated_at = now(),
                      updated_by = 'ingest'
                    """,
                    (
                        project_id,
                        compress_json(bd_raw),
                        compress_json(syn_raw),
                    ),
                )
    print("Done.")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-id", default="default")
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=settings.local_data_path,
    )
    parser.add_argument("--snapshots-only", action="store_true")
    args = parser.parse_args()
    if not settings.use_postgres_store:
        raise SystemExit(
            "Set DATA_BACKEND=postgres (or supabase) and SUPABASE_DB_URL in api/.env"
        )
    ingest(args.project_id, args.data_dir, snapshots_only=args.snapshots_only)


if __name__ == "__main__":
    main()
