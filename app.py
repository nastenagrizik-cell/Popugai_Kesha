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

from crosstab_logic import CrosstabConfig, compute_crosstab, dimension_values, load_survey_excel

app = FastAPI(title="Survey crosstabs")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC = Path(__file__).resolve().parent / "static"
DATASETS: dict[str, Path] = {}


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
    significance_mode: str = "none"
    sig_level: int = 95


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
    columns = [{"key": k, "label": labels.get(k, k)} for k in df.columns]

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


@app.post("/api/crosstab")
def api_crosstab(body: CrosstabBody):
    path = DATASETS.get(body.dataset_id)
    if not path or not path.exists():
        raise HTTPException(404, "Сессия истекла или файл не найден. Загрузите снова.")

    if not body.row_cols or not body.col_cols:
        raise HTTPException(400, "Укажите хотя бы один столбец для строк и для столбцов.")

    if body.row_single and len(body.row_cols) != 1:
        raise HTTPException(400, "Режим «один столбец в строках» — выберите ровно один столбец.")

    if body.col_single and len(body.col_cols) != 1:
        raise HTTPException(400, "Режим «один столбец в столбцах» — выберите ровно один столбец.")

    df, labels = load_survey_excel(str(path))

    cfg = CrosstabConfig(
        row_cols=body.row_cols,
        col_cols=body.col_cols,
        row_single=body.row_single,
        col_single=body.col_single,
        row_include_ids=body.row_include_ids,
        col_include_ids=body.col_include_ids,
        row_scale=body.row_scale,
        include_top2=body.include_top2,
        include_bottom2=body.include_bottom2,
        show_full_scale=body.show_full_scale,
        show_base=body.show_base,
        significance_mode=body.significance_mode,
        sig_level=body.sig_level,
    )

    try:
        result = compute_crosstab(df, labels, cfg)
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
