"""Ingest exported JSON into Delta tables for progressive load."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from app.config import settings
from app.databricks_store import databricks_connection, _table
from app.snapshot_codec import compress_json


def ingest(project_id: str, data_dir: Path, *, snapshots_only: bool = False) -> None:
    files = {
        "BD": data_dir / "bd-sheet.json",
        "SYNTHESIS": data_dir / "synthesis-sheet.json",
    }
    with databricks_connection() as conn:
        with conn.cursor() as cur:
            for sheet, path in files.items():
                if not path.is_file():
                    print(f"Skip missing {path}")
                    continue
                print(f"Loading {path.name} …")
                with path.open(encoding="utf-8") as f:
                    raw = json.load(f)
                meta = {k: v for k, v in raw.items() if k != "cells"}
                meta_json = json.dumps(meta, ensure_ascii=False)
                cur.execute(
                    f"""
                    MERGE INTO {_table('sheet_meta')} t
                    USING (SELECT ? pid, ? sh, ? mj, current_timestamp() ts) s
                    ON t.project_id = s.pid AND t.sheet = s.sh
                    WHEN MATCHED THEN UPDATE SET meta_json = s.mj, updated_at = s.ts
                    WHEN NOT MATCHED THEN INSERT (project_id, sheet, meta_json, updated_at)
                      VALUES (s.pid, s.sh, s.mj, s.ts)
                    """,
                    (project_id, sheet, meta_json),
                )
                if snapshots_only:
                    print(f"  {sheet}: meta only ({len(raw.get('cells') or [])} cells in snapshot)")
                    continue
                cur.execute(
                    f"DELETE FROM {_table('sheet_cells')} WHERE project_id = ? AND sheet = ?",
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
                        cur.executemany(
                            f"""
                            INSERT INTO {_table('sheet_cells')}
                            (project_id, sheet, row, col, v, f, user_edited)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                            """,
                            batch,
                        )
                        batch.clear()
                if batch:
                    cur.executemany(
                        f"""
                        INSERT INTO {_table('sheet_cells')}
                        (project_id, sheet, row, col, v, f, user_edited)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        batch,
                    )
                print(f"  {sheet}: {len(raw.get('cells') or [])} cells")
            bd_path = files["BD"]
            syn_path = files["SYNTHESIS"]
            if not bd_path.is_file() or not syn_path.is_file():
                raise SystemExit("bd-sheet.json and synthesis-sheet.json required")
            bd_snap = compress_json(json.loads(bd_path.read_text(encoding="utf-8")))
            syn_snap = compress_json(json.loads(syn_path.read_text(encoding="utf-8")))
            print("Saving workbook_sessions (compressed BD) …")
            cur.execute(
                f"""
                MERGE INTO {_table('workbook_sessions')} t
                USING (
                  SELECT ? pid, 0 rev, ? bd, CAST(NULL AS STRING) syn,
                         current_timestamp() ts, 'ingest' by
                ) s
                ON t.project_id = s.pid
                WHEN MATCHED THEN UPDATE SET
                  bd_snapshot = s.bd, revision = s.rev, updated_at = s.ts, updated_by = s.by
                WHEN NOT MATCHED THEN INSERT
                  (project_id, revision, bd_snapshot, syn_snapshot, updated_at, updated_by)
                  VALUES (s.pid, s.rev, s.bd, s.syn, s.ts, s.by)
                """,
                (project_id, bd_snap),
            )
            print("Saving workbook_sessions (compressed SYNTHESIS) …")
            cur.execute(
                f"""
                UPDATE {_table('workbook_sessions')}
                SET syn_snapshot = ?, updated_at = current_timestamp(), updated_by = 'ingest'
                WHERE project_id = ?
                """,
                (syn_snap, project_id),
            )
    print("Done.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-id", default="default")
    parser.add_argument("--data-dir", type=Path, default=settings.local_data_path)
    parser.add_argument(
        "--snapshots-only",
        action="store_true",
        help="Skip sheet_cells (very slow). Meta + workbook_sessions only — enough for multi-user MVP.",
    )
    args = parser.parse_args()
    if not settings.use_databricks:
        raise SystemExit("Set DATABRICKS_* in api/.env first")
    ingest(args.project_id, args.data_dir, snapshots_only=args.snapshots_only)


if __name__ == "__main__":
    main()
