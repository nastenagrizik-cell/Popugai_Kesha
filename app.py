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
    show_full_scale: 
