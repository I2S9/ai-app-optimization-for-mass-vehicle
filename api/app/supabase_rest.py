"""Supabase via HTTPS (port 443) — works when Postgres port 6543 is blocked."""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import settings
from .snapshot_codec import compress_json, decompress_json

_HEADERS_EXTRA = {
    "Content-Type": "application/json",
    "Prefer": "return=representation",
}


def _client() -> httpx.Client:
    if not settings.supabase_url or not settings.supabase_service_key:
        raise RuntimeError("SUPABASE_URL and SUPABASE_SERVICE_KEY required")
    base = settings.supabase_url.rstrip("/")
    key = settings.supabase_service_key
    return httpx.Client(
        base_url=f"{base}/rest/v1",
        headers={
            "apikey": key,
            "Authorization": f"Bearer {key}",
            **_HEADERS_EXTRA,
        },
        timeout=httpx.Timeout(120.0, connect=30.0),
    )


def _raise(resp: httpx.Response) -> None:
    if resp.is_success:
        return
    raise RuntimeError(f"Supabase REST {resp.status_code}: {resp.text[:500]}")


def fetch_session(project_id: str) -> dict[str, Any] | None:
    with _client() as c:
        r = c.get(
            "/workbook_sessions",
            params={
                "project_id": f"eq.{project_id}",
                "select": "revision,structure_revision,bd_snapshot,syn_snapshot,updated_at,updated_by",
                "limit": 1,
            },
        )
        _raise(r)
        rows = r.json()
        if not rows:
            return None
        row = rows[0]
        bd = decompress_json(row.get("bd_snapshot"))
        syn = decompress_json(row.get("syn_snapshot"))
        struct_rev = int(row.get("structure_revision") or 0)
        if bd and isinstance(bd, dict) and not struct_rev:
            struct_rev = int(bd.get("structureRevision") or 0)
        return {
            "project_id": project_id,
            "revision": int(row.get("revision") or 0),
            "structureRevision": struct_rev,
            "bd": bd,
            "syn": syn,
            "updated_at": row.get("updated_at"),
            "updated_by": row.get("updated_by"),
        }


def upsert_session(
    project_id: str,
    revision: int,
    bd: dict | None,
    syn: dict | None,
    updated_by: str = "web",
    structure_revision: int = 0,
) -> dict[str, Any]:
    if bd and structure_revision:
        bd = {**bd, "structureRevision": structure_revision}
    if syn and structure_revision:
        syn = {**syn, "structureRevision": structure_revision}

    with _client() as c:
        r = c.get(
            "/workbook_sessions",
            params={"project_id": f"eq.{project_id}", "select": "revision", "limit": 1},
        )
        _raise(r)
        existing = r.json()
        if existing and int(existing[0].get("revision") or 0) > revision:
            return {
                "project_id": project_id,
                "revision": int(existing[0]["revision"]),
                "ok": False,
                "conflict": True,
            }

        payload = {
            "project_id": project_id,
            "revision": revision,
            "structure_revision": structure_revision,
            "bd_snapshot": compress_json(bd) if bd else None,
            "syn_snapshot": compress_json(syn) if syn else None,
            "updated_by": updated_by,
        }
        r2 = c.post(
            "/workbook_sessions",
            headers={**_HEADERS_EXTRA, "Prefer": "resolution=merge-duplicates,return=representation"},
            json=payload,
        )
        _raise(r2)

    return {
        "project_id": project_id,
        "revision": revision,
        "structureRevision": structure_revision,
        "ok": True,
    }


def fetch_meta(project_id: str, sheet: str) -> dict[str, Any] | None:
    with _client() as c:
        r = c.get(
            "/sheet_meta",
            params={
                "project_id": f"eq.{project_id}",
                "sheet": f"eq.{sheet.upper()}",
                "select": "meta_json",
                "limit": 1,
            },
        )
        _raise(r)
        rows = r.json()
        if not rows:
            return None
        data = rows[0].get("meta_json")
        if isinstance(data, str):
            return json.loads(data)
        return dict(data) if data else None


def fetch_cells(
    project_id: str,
    sheet: str,
    row_min: int,
    row_max: int,
) -> list[dict[str, Any]]:
    with _client() as c:
        r = c.get(
            "/sheet_cells",
            params={
                "project_id": f"eq.{project_id}",
                "sheet": f"eq.{sheet.upper()}",
                "and": f"(excel_row.gte.{row_min},excel_row.lte.{row_max})",
                "select": "excel_row,col,v,f,user_edited",
                "order": "excel_row.asc,col.asc",
            },
        )
        _raise(r)
        rows = r.json()
    cells = []
    for row in rows:
        er = int(row["excel_row"])
        cell: dict[str, Any] = {"r": er, "c": str(row["col"])}
        if row.get("v") is not None:
            cell["v"] = str(row["v"])
        if row.get("f"):
            cell["f"] = str(row["f"])
        if row.get("user_edited"):
            cell["userEdited"] = True
        cells.append(cell)
    return cells


def upsert_cells(
    project_id: str,
    sheet: str,
    changes: list[dict[str, Any]],
) -> None:
    if not changes:
        return
    batch = []
    sh = sheet.upper()
    for ch in changes:
        batch.append(
            {
                "project_id": project_id,
                "sheet": sh,
                "excel_row": int(ch["r"]),
                "col": str(ch["c"]),
                "v": None if ch.get("v") is None else str(ch.get("v", "")),
                "f": str(ch["f"]) if ch.get("f") else None,
                "user_edited": True,
            }
        )
    with _client() as c:
        for i in range(0, len(batch), 400):
            chunk = batch[i : i + 400]
            r = c.post(
                "/sheet_cells",
                headers={
                    **_HEADERS_EXTRA,
                    "Prefer": "resolution=merge-duplicates,return=minimal",
                },
                json=chunk,
            )
            _raise(r)


def upsert_meta(project_id: str, sheet: str, meta: dict[str, Any]) -> None:
    with _client() as c:
        r = c.post(
            "/sheet_meta",
            headers={**_HEADERS_EXTRA, "Prefer": "resolution=merge-duplicates,return=minimal"},
            json={
                "project_id": project_id,
                "sheet": sheet.upper(),
                "meta_json": meta,
            },
        )
        _raise(r)


def delete_cells_for_sheet(project_id: str, sheet: str) -> None:
    with _client() as c:
        r = c.delete(
            "/sheet_cells",
            params={"project_id": f"eq.{project_id}", "sheet": f"eq.{sheet.upper()}"},
        )
        _raise(r)


def fetch_grid_pack(project_id: str, sheet: str) -> dict[str, Any] | None:
    """Return the precomputed display grid pack ({fingerprint, sheet, ...}) or None."""
    with _client() as c:
        r = c.get(
            "/sheet_grid_cache",
            params={
                "project_id": f"eq.{project_id}",
                "sheet": f"eq.{sheet.upper()}",
                "select": "fingerprint,pack",
                "limit": 1,
            },
        )
        _raise(r)
        rows = r.json()
        if not rows:
            return None
        pack = decompress_json(rows[0].get("pack"))
        if not isinstance(pack, dict):
            return None
        if not pack.get("fingerprint"):
            pack["fingerprint"] = rows[0].get("fingerprint") or ""
        return pack


def upsert_grid_pack(
    project_id: str,
    sheet: str,
    fingerprint: str,
    pack: dict[str, Any],
) -> None:
    """Store the compressed display grid pack for a sheet (one row per project/sheet)."""
    with _client() as c:
        r = c.post(
            "/sheet_grid_cache",
            headers={**_HEADERS_EXTRA, "Prefer": "resolution=merge-duplicates,return=minimal"},
            json={
                "project_id": project_id,
                "sheet": sheet.upper(),
                "fingerprint": fingerprint,
                "pack": compress_json(pack),
            },
        )
        _raise(r)


def ingest_sheet(
    project_id: str,
    sheet: str,
    raw: dict[str, Any],
    progress: Any | None = None,
) -> int:
    meta = {k: v for k, v in raw.items() if k != "cells"}
    print(f"    meta {sheet}…", flush=True)
    upsert_meta(project_id, sheet, meta)
    print(f"    clear old cells {sheet}…", flush=True)
    delete_cells_for_sheet(project_id, sheet)
    cells_in = raw.get("cells") or []
    batch = []
    count = 0
    for cell in cells_in:
        batch.append(
            {
                "project_id": project_id,
                "sheet": sheet.upper(),
                "excel_row": int(cell["r"]),
                "col": str(cell["c"]),
                "v": None if cell.get("v") is None else str(cell.get("v", "")),
                "f": str(cell["f"]) if cell.get("f") else None,
                "user_edited": bool(cell.get("userEdited")),
            }
        )
        if len(batch) >= 400:
            with _client() as c:
                r = c.post(
                    "/sheet_cells",
                    headers={
                        **_HEADERS_EXTRA,
                        "Prefer": "resolution=merge-duplicates,return=minimal",
                    },
                    json=batch,
                )
                _raise(r)
            count += len(batch)
            if progress:
                progress(count)
            batch.clear()
    if batch:
        with _client() as c:
            r = c.post(
                "/sheet_cells",
                headers={
                    **_HEADERS_EXTRA,
                    "Prefer": "resolution=merge-duplicates,return=minimal",
                },
                json=batch,
            )
            _raise(r)
        count += len(batch)
        if progress:
            progress(count)
    return count
