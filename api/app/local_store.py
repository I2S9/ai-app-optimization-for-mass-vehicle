"""Local JSON sheet store — same chunk API as Databricks target."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

SHEET_FILES = {
    "bd": "bd-sheet.json",
    "synthesis": "synthesis-sheet.json",
}


class LocalSheetStore:
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self._cache: dict[str, dict[str, Any]] = {}

    def _load(self, sheet_id: str) -> dict[str, Any]:
        if sheet_id in self._cache:
            return self._cache[sheet_id]
        fname = SHEET_FILES.get(sheet_id)
        if not fname:
            raise KeyError(f"Unknown sheet: {sheet_id}")
        path = self.data_dir / fname
        if not path.is_file():
            raise FileNotFoundError(f"Missing {path}")
        with path.open(encoding="utf-8") as f:
            data = json.load(f)
        self._cache[sheet_id] = data
        return data

    def get_meta(self, sheet_id: str) -> dict[str, Any]:
        raw = self._load(sheet_id)
        meta = {k: v for k, v in raw.items() if k != "cells"}
        meta["cellCount"] = len(raw.get("cells") or [])
        return meta

    def get_cells(self, sheet_id: str, row_min: int, row_max: int) -> list[dict]:
        raw = self._load(sheet_id)
        cells = []
        for cell in raw.get("cells") or []:
            r = int(cell.get("r", 0))
            if row_min <= r <= row_max:
                cells.append(cell)
        return cells

    def get_full(self, sheet_id: str) -> dict[str, Any]:
        return self._load(sheet_id)
