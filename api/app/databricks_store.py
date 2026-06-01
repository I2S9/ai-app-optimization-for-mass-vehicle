"""Databricks SQL warehouse access."""

from __future__ import annotations

import json
from contextlib import contextmanager
from typing import Any, Iterator

from .config import settings
from .snapshot_codec import compress_json, decompress_json


@contextmanager
def databricks_connection():
    from databricks import sql

    conn = sql.connect(
        server_hostname=settings.databricks_host.replace("https://", "").replace("http://", ""),
        http_path=settings.databricks_http_path,
        access_token=settings.databricks_token,
    )
    try:
        yield conn
    finally:
        conn.close()


def _table(name: str) -> str:
    return f"{settings.databricks_catalog}.{settings.databricks_schema}.{name}"


def fetch_session(project_id: str) -> dict[str, Any] | None:
    q = f"""
        SELECT revision, bd_snapshot, syn_snapshot, updated_at, updated_by
        FROM {_table('workbook_sessions')}
        WHERE project_id = ?
        ORDER BY revision DESC
        LIMIT 1
    """
    with databricks_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(q, (project_id,))
            row = cur.fetchone()
            if not row:
                return None
            revision, bd_snap, syn_snap, updated_at, updated_by = row
            return {
                "project_id": project_id,
                "revision": int(revision or 0),
                "bd": decompress_json(bd_snap),
                "syn": decompress_json(syn_snap),
                "updated_at": str(updated_at) if updated_at else None,
                "updated_by": updated_by,
            }


def upsert_session(
    project_id: str,
    revision: int,
    bd: dict | None,
    syn: dict | None,
    updated_by: str = "api",
) -> dict[str, Any]:
    q = f"""
        MERGE INTO {_table('workbook_sessions')} AS t
        USING (
          SELECT ? AS project_id, ? AS revision,
                 ? AS bd_snapshot, ? AS syn_snapshot,
                 current_timestamp() AS updated_at, ? AS updated_by
        ) AS s
        ON t.project_id = s.project_id
        WHEN MATCHED AND t.revision <= s.revision THEN UPDATE SET
          revision = s.revision,
          bd_snapshot = s.bd_snapshot,
          syn_snapshot = s.syn_snapshot,
          updated_at = s.updated_at,
          updated_by = s.updated_by
        WHEN NOT MATCHED THEN INSERT *
    """
    bd_json = compress_json(bd) if bd else None
    syn_json = compress_json(syn) if syn else None
    with databricks_connection() as conn:
        with conn.cursor() as cur:
            if bd_json and syn_json and len(bd_json) + len(syn_json) > 1_900_000:
                cur.execute(
                    q,
                    (project_id, revision, bd_json, None, updated_by),
                )
                cur.execute(
                    f"""
                    UPDATE {_table('workbook_sessions')}
                    SET syn_snapshot = ?, updated_at = current_timestamp(), updated_by = ?
                    WHERE project_id = ? AND revision <= ?
                    """,
                    (syn_json, updated_by, project_id, revision),
                )
            else:
                cur.execute(
                    q,
                    (project_id, revision, bd_json, syn_json, updated_by),
                )
    return {"project_id": project_id, "revision": revision, "ok": True}


def fetch_meta(project_id: str, sheet: str) -> dict[str, Any] | None:
    q = f"""
        SELECT meta_json FROM {_table('sheet_meta')}
        WHERE project_id = ? AND sheet = ?
        LIMIT 1
    """
    with databricks_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(q, (project_id, sheet.upper()))
            row = cur.fetchone()
            if not row or not row[0]:
                return None
            return json.loads(row[0])


def fetch_cells(
    project_id: str,
    sheet: str,
    row_min: int,
    row_max: int,
) -> list[dict[str, Any]]:
    q = f"""
        SELECT row, col, v, f, user_edited
        FROM {_table('sheet_cells')}
        WHERE project_id = ? AND sheet = ?
          AND row BETWEEN ? AND ?
    """
    with databricks_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(q, (project_id, sheet.upper(), row_min, row_max))
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
