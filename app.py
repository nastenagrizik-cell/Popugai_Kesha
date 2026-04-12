from __future__ import annotations

import os
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from crosstab_logic import (
    CrosstabConfig,
    compute_crosstab,
    dimension_values,
    get_question_coding_candidates,
    is_question_coding_candidate,
    load_survey_excel,
)

app = FastAPI(title="Survey crosstabs")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC = Path(__file__).resolve().parent / "static"
DATASETS: dict[str, Path] = {}
MANUAL_CODINGS: dict[str, dict[str, dict[str, int | None]]] = {}


class DimensionsBody(BaseModel):
    dataset_id: str
    cols: list[str]
    single: bool


class CrosstabBody(BaseModel):
    dataset_id: str
    row_cols: list[str]
    col_cols: list[str]
    row_single: bool = False
    col_single: bool = False
    row_include_ids: list[str] | None = None
    col_include_ids: list[str] | None = None
    row_scale: bool = False
    include_top2: bool = False
    include_bottom2: bool = False
    show_full_scale: bool = True
    show_base: bool = True
    significance_mode: str = "none"  # none | vs_first | all_vs_all
    sig_level: int = 95  # 90 | 95


class CodingPreviewBody(BaseModel):
    dataset_id: str
    question_key: str


class CodingSaveBody(BaseModel):
    dataset_id: str
    question_key: str
    mapping: dict[str, int | None]


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Нужен файл .xlsx или .xls")

    did = uuid.uuid4().hex
    suffix = ".xlsx" if file.filename.lower().endswith(".xlsx") else ".xls"
    path = Path(tempfile.gettempdir()) / f"survey_{did}{suffix}"

    content = await file.read()
    if len(content) > 80 * 1024 * 1024:
        raise HTTPException(400, "Файл больше 80 МБ")

    path.write_bytes(content)

    try:
        df, labels = load_survey_excel(str(path))
    except Exception as e:
        path.unlink(missing_ok=True)
        raise HTTPException(400, str(e)) from e

    DATASETS[did] = path
    MANUAL_CODINGS[did] = {}

    columns = []
    for k in df.columns:
        columns.append(
            {
                "key": k,
                "label": labels.get(k, k),
                "can_code": is_question_coding_candidate(df[k]),
            }
        )

    return {
        "dataset_id": did,
        "n_rows": int(len(df)),
        "columns": columns,
    }


@app.post("/api/dimensions")
def api_dimensions(body: DimensionsBody):
    path = DATASETS.get(body.dataset_id)
    if not path or not path.exists():
        raise HTTPException(404, "Сессия истекла или файл не найден. Загрузите снова.")

    df, lbl = load_survey_excel(str(path))

    for c in body.cols:
        if c not in df.columns:
            raise HTTPException(400, f"Неизвестный столбец: {c}")

    try:
        items = dimension_values(df, body.cols, single=body.single, labels=lbl)
    except Exception as e:
        raise HTTPException(400, str(e)) from e

    return {"values": items}


@app.post("/api/coding-preview")
def api_coding_preview(body: CodingPreviewBody):
    path = DATASETS.get(body.dataset_id)
    if not path or not path.exists():
        raise HTTPException(404, "Сессия истекла или файл не найден. Загрузите снова.")

    df, labels = load_survey_excel(str(path))

    if body.question_key not in df.columns:
        raise HTTPException(400, f"Неизвестный столбец: {body.question_key}")

    series = df[body.question_key]
    if not is_question_coding_candidate(series):
        raise HTTPException(
            400,
            "Этот вопрос не похож на закрытый/шкальный. Кодировка для него скрыта, чтобы не трогать открытые вопросы.",
        )

    existing = MANUAL_CODINGS.get(body.dataset_id, {}).get(body.question_key, {})
    items = get_question_coding_candidates(series, existing)

    return {
        "question_key": body.question_key,
        "question_label": labels.get(body.question_key, body.question_key),
        "items": items,
    }


@app.post("/api/coding-save")
def api_coding_save(body: CodingSaveBody):
    path = DATASETS.get(body.dataset_id)
    if not path or not path.exists():
        raise HTTPException(404, "Сессия истекла или файл не найден. Загрузите снова.")

    df, _labels = load_survey_excel(str(path))
    if body.question_key not in df.columns:
        raise HTTPException(400, f"Неизвестный столбец: {body.question_key}")

    if body.dataset_id not in MANUAL_CODINGS:
        MANUAL_CODINGS[body.dataset_id] = {}

    cleaned: dict[str, int | None] = {}
    for raw, code in body.mapping.items():
        raw2 = str(raw).strip()
        if not raw2:
            continue
        if code is None:
            cleaned[raw2] = None
        elif code in (1, 2, 3, 4, 5):
            cleaned[raw2] = int(code)
        else:
            raise HTTPException(400, f"Недопустимый код для значения '{raw2}': {code}")

    MANUAL_CODINGS[body.dataset_id][body.question_key] = cleaned
    return {"ok": True, "saved": len(cleaned)}


@app.post("/api/crosstab")
def api_crosstab(body: CrosstabBody):
    path = DATASETS.get(body.dataset_id)
    if not path or not path.exists():
        raise HTTPException(404, "Сессия истекла или файл не найден. Загрузите снова.")

    if not body.row_cols:
        raise HTTPException(400, "Укажите хотя бы один столбец для строк.")

    if body.significance_mode not in {"none", "vs_first", "all_vs_all"}:
        raise HTTPException(400, "Недопустимый режим значимости")
    if body.sig_level not in {90, 95}:
        raise HTTPException(400, "Уровень значимости может быть только 90 или 95")

    df, labels = load_survey_excel(str(path))

    for c in body.row_cols + body.col_cols:
        if c not in df.columns:
            raise HTTPException(400, f"Неизвестный столбец: {c}")

    manual_codings = MANUAL_CODINGS.get(body.dataset_id, {})

    try:
        cfg = CrosstabConfig(
            row_cols=body.row_cols,
            col_cols=body.col_cols,
            include_top2=body.include_top2,
            include_bottom2=body.include_bottom2,
            show_full_scale=body.show_full_scale,
            show_base=body.show_base,
            significance_mode=body.significance_mode,
            sig_level=body.sig_level,
        )
        result = compute_crosstab(df, labels, cfg, manual_codings=manual_codings)
    except Exception as e:
        raise HTTPException(400, str(e)) from e

    return result


@app.get("/api/health")
def health():
    return {"ok": True}


if STATIC.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC)), name="static")


@app.get("/")
def root():
    index = STATIC / "index.html"
    if index.exists():
        return FileResponse(index)
    return {"message": "Place static/index.html"}


def main():
    import uvicorn

    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run("app:app", host="0.0.0.0", port=port)


if __name__ == "__main__":
    main()
