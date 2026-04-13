from __future__ import annotations

import math
import re
from collections import Counter
from dataclasses import dataclass
from typing import Any

import pandas as pd

LOW_BASE_THRESHOLD = 30


def _cell_nonempty(v: Any) -> bool:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    s = str(v).strip()
    if not s:
        return False
    low = s.lower()
    if low in ("0", "false", "нет", "no", "nan", "none"):
        return False
    return True


def _normalize_text(v: Any) -> str:
    s = str(v).strip().lower()
    s = s.replace("ё", "е")
    s 
