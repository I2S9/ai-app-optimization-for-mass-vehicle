"""Supabase / PostgreSQL — same contract as databricks_store."""

from __future__ import annotations

import json
from contextlib import contextmanager
from typing import Any, Iterator

from .config import settings
from .snapshot_codec import compress_json, decompress_json

SHEET_BD = "BD"
SHEET_SYN = "SYNTHESIS"


@contextmanager
def pg_connection():
    import psycopg2

    if not settings.pg_url:
        raise RuntimeError("SUPABASE_DB_URL is not set")
    conn = psycopg2.connect(settings.pg_url, connect_timeout=8)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _stamp_structure_revision(sheet: dict | None, structure_revision: int) -> dict | None:
    if not sheet or not structure_revision:
        return sheet
    out = dict(sheet)
    out["structureRevision"] = structure_revision
    return out


def fetch_session(project_id: str) -> dict[str, Any] | None:
    with pg_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT revision, structure_revision, bd_snapshot, syn_snapshot,
                       updated_at, updated_by
                FROM public.workbook_sessions
                WHERE project_id = %s
                """,
                (project_id,),
            )
            row = cur.fetchone()
            if not row:
                return None
            revision, struct_rev, bd_snap, syn_snap, updated_at, updated_by = row
            bd = decompress_json(bd_snap)
            syn = decompress_json(syn_snap)
            struct_rev = int(struct_rev or 0)
            if bd and isinstance(bd, dict) and not struct_rev:
                struct_rev = int(bd.get("structureRevision") or 0)
            return {
                "project_id": project_id,
                "revision": int(revision or 0),
                "structureRevision": struct_rev,
                "bd": bd,
                "syn": syn,
                "updated_at": updated_at.isoformat() if updated_at else None,
                "updated_by": updated_by,
            }


def upsert_session(
    project_id: str,
    revision: int,
    bd: dict | None,
    syn: dict | None,
    updated_by: str = "web",
    structure_revision: int = 0,
) -> dict[str, Any]:
    bd_stamped = _stamp_structure_revision(bd, structure_revision)
    syn_stamped = _stamp_structure_revision(syn, structure_revision)
    bd_json = compress_json(bd_stamped) if bd_stamped else None
    syn_json = compress_json(syn_stamped) if syn_stamped else None

    with pg_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT revision FROM public.workbook_sessions WHERE project_id = %s",
                (project_id,),
            )
            existing = cur.fetchone()
            if existing and int(existing[0] or 0) > revision:
                return {
                    "project_id": project_id,
                    "revision": int(existing[0]),
                    "ok": False,
                    "conflict": True,
                }

            cur.execute(
                """
                INSERT INTO public.workbook_sessions (
                  project_id, revision, structure_revision,
                  bd_snapshot, syn_snapshot, updated_at, updated_by
                )
                VALUES (%s, %s, %s, %s, %s, now(), %s)
                ON CONFLICT (project_id) DO UPDATE SET
                  revision = EXCLUDED.revision,
                  structure_revision = EXCLUDED.structure_revision,
                  bd_snapshot = EXCLUDED.bd_snapshot,
                  syn_snapshot = EXCLUDED.syn_snapshot,
                  updated_at = now(),
                  updated_by = EXCLUDED.updated_by
                WHERE public.workbook_sessions.revision <= EXCLUDED.revision
                """,
                (
                    project_id,
                    revision,
                    structure_revision,
                    bd_json,
                    syn_json,
                    updated_by,
                ),
            )

            if syn_json and bd_json and len(bd_json) + len(syn_json) > 1_900_000:
                cur.execute(
                    """
                    UPDATE public.workbook_sessions
                    SET syn_snapshot = %s, updated_at = now(), updated_by = %s
                    WHERE project_id = %s AND revision <= %s
                    """,
                    (syn_json, updated_by, project_id, revision),
                )

    return {
        "project_id": project_id,
        "revision": revision,
        "structureRevision": structure_revision,
        "ok": True,
    }


def fetch_meta(project_id: str, sheet: str) -> dict[str, Any] | None:
    with pg_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT meta_json FROM public.sheet_meta
                WHERE project_id = %s AND sheet = %s
                """,
                (project_id, sheet.upper()),
            )
            row = cur.fetchone()
            if not row or row[0] is None:
                return None
            data = row[0]
            if isinstance(data, str):
                return json.loads(data)
            return dict(data)


def fetch_cells(
    project_id: str,
    sheet: str,
    row_min: int,
    row_max: int,
) -> list[dict[str, Any]]:
    with pg_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT excel_row, col, v, f, user_edited
                FROM public.sheet_cells
                WHERE project_id = %s AND sheet = %s
                  AND excel_row BETWEEN %s AND %s
                ORDER BY excel_row, col
                """,
                (project_id, sheet.upper(), row_min, row_max),
            )
            rows = cur.fetchall()
    cells = []
    for r, col, v, f, user_edited in rows:
        cell: dict[str, Any] = {"r": int(r), "c": str(col)}
        if v is not None:
            cell["v"] = str(v)
        if f:
            cell["f"] = str(f)
        if user_edited:
            cell["userEdited"] = True
        cells.append(cell)
    return cells


@contextmanager
def ingest_connection() -> Iterator[Any]:
    """Long-running ingest — caller commits."""
    import psycopg2

    conn = psycopg2.connect(settings.pg_url, connect_timeout=8)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def fetch_module_state(project_id: str, module_key: str) -> dict[str, Any] | None:
    with pg_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT revision, state_json, updated_at, updated_by
                FROM public.module_state
                WHERE project_id = %s AND module_key = %s
                """,
                (project_id, module_key),
            )
            row = cur.fetchone()
            if not row:
                return None
            revision, state_json, updated_at, updated_by = row
            state = state_json
            if isinstance(state, str):
                state = json.loads(state)
            return {
                "project_id": project_id,
                "module_key": module_key,
                "revision": int(revision or 0),
                "state": dict(state) if state else {},
                "updated_at": updated_at.isoformat() if updated_at else None,
                "updated_by": updated_by,
            }


def upsert_module_state(
    project_id: str,
    module_key: str,
    revision: int,
    state: dict[str, Any],
    updated_by: str = "web",
) -> dict[str, Any]:
    with pg_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT revision FROM public.module_state
                WHERE project_id = %s AND module_key = %s
                """,
                (project_id, module_key),
            )
            existing = cur.fetchone()
            if existing and int(existing[0] or 0) > revision:
                return {
                    "project_id": project_id,
                    "module_key": module_key,
                    "revision": int(existing[0]),
                    "ok": False,
                    "conflict": True,
                }

            cur.execute(
                """
                INSERT INTO public.module_state (
                  project_id, module_key, revision, state_json, updated_at, updated_by
                ) VALUES (%s, %s, %s, %s, now(), %s)
                ON CONFLICT (project_id, module_key) DO UPDATE SET
                  revision = EXCLUDED.revision,
                  state_json = EXCLUDED.state_json,
                  updated_at = now(),
                  updated_by = EXCLUDED.updated_by
                WHERE public.module_state.revision <= EXCLUDED.revision
                RETURNING revision
                """,
                (project_id, module_key, revision, json.dumps(state), updated_by),
            )
            saved = cur.fetchone()

    saved_rev = int(saved[0]) if saved else revision
    return {
        "project_id": project_id,
        "module_key": module_key,
        "revision": saved_rev,
        "ok": True,
    }
