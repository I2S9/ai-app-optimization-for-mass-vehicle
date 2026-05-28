"""FastAPI — progressive sheet load + Databricks session snapshots."""

from __future__ import annotations

from typing import Any

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from .config import settings
from .local_store import LocalSheetStore

app = FastAPI(title="WGHT Vehicle Mass API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origin_list,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

local_store = LocalSheetStore(settings.local_data_path)
DEFAULT_PROJECT = "default"


def _sheet_key(sheet_id: str) -> str:
    s = sheet_id.lower()
    if s not in ("bd", "synthesis"):
        raise HTTPException(404, f"Unknown sheet '{sheet_id}'")
    return s


def _databricks_sheet_name(sheet_id: str) -> str:
    return "BD" if sheet_id == "bd" else "SYNTHESIS"


@app.get("/api/v1/health")
def health():
    return {
        "ok": True,
        "backend": "databricks" if settings.use_databricks else "local",
    }


@app.get("/api/v1/config")
def config():
    return {
        "mode": "databricks" if settings.use_databricks else "local",
        "chunkedLoad": True,
        "version": 1,
        "projectId": DEFAULT_PROJECT,
    }


@app.get("/api/v1/sheets/{sheet_id}/meta")
def sheet_meta(sheet_id: str, project_id: str = DEFAULT_PROJECT):
    sheet_id = _sheet_key(sheet_id)
    if settings.use_databricks:
        try:
            from . import databricks_store

            meta = databricks_store.fetch_meta(project_id, _databricks_sheet_name(sheet_id))
            if meta:
                return meta
        except Exception as exc:
            raise HTTPException(503, f"Databricks meta: {exc}") from exc
    try:
        return local_store.get_meta(sheet_id)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc


@app.get("/api/v1/sheets/{sheet_id}/cells")
def sheet_cells(
    sheet_id: str,
    rowMin: int = Query(..., ge=1),
    rowMax: int = Query(..., ge=1),
    project_id: str = DEFAULT_PROJECT,
):
    if rowMax < rowMin:
        raise HTTPException(400, "rowMax must be >= rowMin")
    sheet_id = _sheet_key(sheet_id)
    if settings.use_databricks:
        try:
            from . import databricks_store

            cells = databricks_store.fetch_cells(
                project_id,
                _databricks_sheet_name(sheet_id),
                rowMin,
                rowMax,
            )
            if cells:
                return {"cells": cells, "rowMin": rowMin, "rowMax": rowMax}
        except Exception as exc:
            raise HTTPException(503, f"Databricks cells: {exc}") from exc
    try:
        cells = local_store.get_cells(sheet_id, rowMin, rowMax)
        return {"cells": cells, "rowMin": rowMin, "rowMax": rowMax}
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc


class SessionPutBody(BaseModel):
    revision: int = 0
    bd: dict[str, Any] | None = None
    syn: dict[str, Any] | None = None
    updated_by: str = "web"


@app.get("/api/v1/sessions/{project_id}")
def get_session(project_id: str):
    if not settings.use_databricks:
        raise HTTPException(501, "Session API requires Databricks (set DATA_BACKEND=databricks)")
    try:
        from . import databricks_store

        session = databricks_store.fetch_session(project_id)
        if not session:
            raise HTTPException(404, "Session not found")
        return session
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(503, str(exc)) from exc


@app.put("/api/v1/sessions/{project_id}")
def put_session(project_id: str, body: SessionPutBody):
    if not settings.use_databricks:
        raise HTTPException(501, "Session API requires Databricks")
    try:
        from . import databricks_store

        return databricks_store.upsert_session(
            project_id,
            body.revision,
            body.bd,
            body.syn,
            body.updated_by,
        )
    except Exception as exc:
        raise HTTPException(503, str(exc)) from exc
