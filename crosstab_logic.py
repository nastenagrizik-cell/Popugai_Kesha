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


def dimension_values(df: pd.DataFrame, cols: list[str], *, single: bool, labels: dict[str, str] | None = None) -> list[dict[str, Any]]:
    labels = labels or {}
    items: list[dict[str, Any]] = []
    if single and cols:
        col = cols[0]
        vals = []
        for v in df[col].dropna().unique():
            t = str(v).strip()
            if t:
                vals.append(t)
        for v in sorted(set(vals), key=lambda x: (len(x), x)):
            n = int(df[col].map(lambda x, vv=v: _cell_nonempty(x) and str(x).strip() == vv).sum())
            items.append({"id": v, "label": v, "n": n})
    else:
        for c in cols:
            n = int(df[c].map(_cell_nonempty).sum())
            items.append({"id": c, "label": labels.get(c, c), "n": n})
    return items


def is_question_coding_candidate(series: pd.Series) -> bool:
    vals = [str(v).strip() for v in series.dropna() if _cell_nonempty(v)]
    if not vals:
        return False
    unique_vals = list(dict.fromkeys(vals))
    n_unique = len(unique_vals)
    if n_unique > 20:
        return False
    avg_len = sum(len(v) for v in vals) / len(vals)
    if avg_len > 45:
        return False
    uniq_ratio = n_unique / max(len(vals), 1)
    if uniq_ratio > 0.7 and n_unique > 8:
        return False
    return True


def get_question_coding_candidates(series: pd.Series, existing_map: dict[str, int | None] | None = None) -> list[dict[str, Any]]:
    existing_map = existing_map or {}
    counter = Counter()
    for v in series.dropna():
        if _cell_nonempty(v):
            counter[str(v).strip()] += 1
    items = []
    for raw, n in counter.most_common():
        suggested = parse_scale_1_5(raw)
        current = existing_map.get(raw, suggested)
        items.append({"raw_value": raw, "n": int(n), "suggested_code": suggested, "current_code": current})
    return items


def detect_scale_labels(series_raw: pd.Series, manual_map: dict[str, int | None] | None = None) -> dict[int, str]:
    found: dict[int, str] = {}
    values = []
    for v in series_raw.dropna().unique():
        if _cell_nonempty(v):
            values.append(str(v).strip())
    for val in values:
        score = parse_scale_1_5(val, manual_map=manual_map)
        if score is not None and score not in found:
            found[score] = val
    defaults = {1: "1", 2: "2", 3: "3", 4: "4", 5: "5"}
    return {k: found.get(k, defaults[k]) for k in [1, 2, 3, 4, 5]}


def _norm_cdf(x: float) -> float:
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))


def _two_prop_pvalue(x1: int, n1: int, x2: int, n2: int) -> float | None:
    if min(n1, n2) < LOW_BASE_THRESHOLD:
        return None
    if n1 <= 0 or n2 <= 0:
        return None
    p1 = x1 / n1
    p2 = x2 / n2
    p = (x1 + x2) / (n1 + n2)
    se = math.sqrt(max(p * (1 - p) * (1 / n1 + 1 / n2), 0.0))
    if se == 0:
        return None
    z = (p1 - p2) / se
    return 2 * (1 - _norm_cdf(abs(z)))


def _sig_alpha(sig_level: int) -> float:
    return 0.10 if sig_level == 90 else 0.05


def _col_letter(idx: int) -> str:
    s = ""
    n = idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


@dataclass
class CrosstabConfig:
    row_cols: list[str]
    col_cols: list[str]
    include_top2: bool = False
    include_bottom2: bool = False
    show_full_scale: bool = True
    show_base: bool = True
    significance_mode: str = "none"
    sig_level: int = 95


def compute_crosstab(df: pd.DataFrame, labels: dict[str, str], cfg: CrosstabConfig, manual_codings: dict[str, dict[str, int | None]] | None = None) -> dict[str, Any]:
    manual_codings = manual_codings or {}
    for c in cfg.row_cols + cfg.col_cols:
        if c not in df.columns:
            raise 
