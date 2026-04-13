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
    s = re.sub(r"\s+", " ", s)
    return s


def _uniquify_keys(keys: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for k in keys:
        base = (k or "").strip() or "var"
        if base not in seen:
            seen[base] = 0
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}__{seen[base]}")
    return out


def load_survey_excel(path: str) -> tuple[pd.DataFrame, dict[str, str]]:
    pl = path.lower()
    engine = "xlrd" if pl.endswith(".xls") and not pl.endswith(".xlsx") else "openpyxl"
    raw = pd.read_excel(path, header=None, dtype=object, engine=engine)

    if raw.shape[0] < 3:
        raise ValueError("В файле должно быть минимум 3 строки: ключи, подписи, данные.")

    row0 = raw.iloc[0].tolist()
    row1 = raw.iloc[1].tolist()
    data = raw.iloc[2:].reset_index(drop=True)

    keys: list[str] = []
    labels: dict[str, str] = {}

    for j in range(raw.shape[1]):
        key = str(row0[j]).strip() if pd.notna(row0[j]) else f"col_{j}"
        keys.append(key)

    keys = _uniquify_keys(keys)

    for j, key in enumerate(keys):
        lab = row1[j] if j < len(row1) else None
        labels[key] = str(lab).strip() if pd.notna(lab) and str(lab).strip() else key

    data.columns = keys
    return data, labels


SCALE_TEXT_TO_NUM: dict[str, int] = {
    "совсем не понятна": 1,
    "скорее не понятна": 2,
    "отчасти понятна, отчасти нет": 3,
    "скорее понятна": 4,
    "абсолютно понятна": 5,
    "совсем не подходит": 1,
    "скорее не подходит": 2,
    "отчасти подходит, отчасти нет": 3,
    "скорее подходит": 4,
    "точно подходит": 5,
    "совсем не легко": 1,
    "не очень легко": 2,
    "отчасти легко, отчасти нет": 3,
    "скорее легко": 4,
    "очень легко": 5,
    "совсем не отличается": 1,
    "скорее не отличается": 2,
    "отчасти отличается, отчасти нет": 3,
    "скорее отличается": 4,
    "очень отличается": 5,
    "точно не воспользуюсь": 1,
    "скорее не воспользуюсь": 2,
    "может воспользуюсь, а может, нет": 3,
    "скорее воспользуюсь": 4,
    "точно воспользуюсь": 5,
    "абсолютно не согласен(-а)": 1,
    "скорее не согласен(-а)": 2,
    "отчасти согласен, отчасти нет": 3,
    "отчасти согласен(-а), отчасти нет": 3,
    "скорее согласен(-а)": 4,
    "абсолютно согласен(-а)": 5,
    "точно не поучаствую": 1,
    "скорее не поучаствую": 2,
    "может поучаствую, а может, нет": 3,
    "скорее поучаствую": 4,
    "точно поучаствую": 5,
    "совсем не запомнится": 1,
    "точно запомнится": 5,
    "1 - совсем не запомнится": 1,
    "5 - точно запомнится": 5,
}


def parse_scale_1_5(v: Any, manual_map: dict[str, int | None] | None = None) -> int | None:
    if not _cell_nonempty(v):
        return None

    s = str(v).strip()
    low = _normalize_text(s)

    if manual_map is not None:
        if s in manual_map:
            return manual_map[s]
        norm_map = {_normalize_text(k): val for k, val in manual_map.items()}
        if low in norm_map:
            return norm_map[low]

    if low in SCALE_TEXT_TO_NUM:
        return SCALE_TEXT_TO_NUM[low]

    m = re.match(r"^\s*([1-5])(?:\D|$)", s)
    if m:
        return int(m.group(1))

    try:
        d = int(float(s))
        if 1 <= d <= 5:
            return d
    except (TypeError, ValueError):
        pass

    hints = [
        (1, ["совсем не", "точно не", "абсолютно не"]),
        (2, ["скорее не", "не очень"]),
        (3, ["отчасти", "может", "частично"]),
        (4, ["скорее ", "скорее"]),
        (5, ["абсолютно", "очень", "точно ", "точно"]),
    ]
    for score, words in hints:
        if any(w in low for w in words):
            return score
    return None


def answered_row_multi(df: pd.DataFrame, row_cols: list[str]) -> pd.Series:
    mask = pd.Series(False, index=df.index)
    for c in row_cols:
        mask = mask | df[c].map(_cell_nonempty)
    return mask


def answered_col_multi(df: pd.DataFrame, col_cols: list[str]) -> pd.Series:
    mask = pd.Series(False, index=df.index)
    for c in col_cols:
        mask = mask | df[c].map(_cell_nonempty)
    return mask


def membership_multi(df: pd.DataFrame, cols: list[str]) -> dict[str, pd.Series]:
    return {c: df[c].map(_cell_nonempty) for c in cols}


def membership_single(df: pd.DataFrame, col: str) -> dict[str, pd.Series]:
    s = df[col]
    out: dict[str, pd.Series] = {}
    vals = []
    for v in s.dropna().unique():
        t = str(v).strip()
        if t:
            vals.append(t)
    for t in sorted(set(vals), key=lambda x: (len(x), x)):
        out[t] = s.map(lambda x, tt=t: _cell_nonempty(x) and str(x).strip() == tt)
    return out


def dimension_values(
    df: pd.DataFrame,
    cols: list[str],
    *,
    single: bool,
    labels: dict[str, str] | None = None,
) -> list[dict[str, 
