"""Compress large JSON snapshots for Databricks SQL (1 MB parameter limit)."""

from __future__ import annotations

import base64
import json
import zlib
from typing import Any

PREFIX = "gz+b64:"


def compress_json(data: Any) -> str:
    raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
    blob = base64.b64encode(zlib.compress(raw, 9)).decode("ascii")
    return f"{PREFIX}{blob}"


def decompress_json(text: str | None) -> Any | None:
    if not text:
        return None
    if text.startswith(PREFIX):
        raw = zlib.decompress(base64.b64decode(text[len(PREFIX) :]))
        return json.loads(raw.decode("utf-8"))
    return json.loads(text)
