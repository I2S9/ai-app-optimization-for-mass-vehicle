"""Allow running scripts as: python tools/xxx.py from the api/ folder."""
from __future__ import annotations

import sys
from pathlib import Path

_API_ROOT = Path(__file__).resolve().parent.parent
_root = str(_API_ROOT)
if _root not in sys.path:
    sys.path.insert(0, _root)
