"""
base_loader.py — Carga y pre-indexado de la MATRIZ MATERIALES BASE.
"""
import unicodedata
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional
import pandas as pd


BASE_PATH = Path(__file__).parent.parent / "assets" / "MATRIZ MATERIALES BASE.xlsx"
SHEET_NAME = "MATRIZ DIG"

# Índices de columna (0-based) en la hoja MATRIZ DIG, headers en fila 2 (índice 1)
# MEDIO=col3(C), FORMATO=col4(D), SPECCS=col5(E), PESO=col6(F),
# EXTENSION=col7(G), TEXTO=col8(H), SEGUNDAJE=col9(I)
COL_MEDIO = 2
COL_FORMATO = 3
COL_SPECCS = 4
COL_PESO = 5
COL_EXTENSION = 6
COL_TEXTO = 7
COL_SEGUNDAJE = 8


@dataclass
class Spec:
    medio: str
    formato: str
    speccs: str = ""
    peso: str = ""
    extension: str = ""
    texto: str = ""
    segundaje_ideal: str = ""
    # Versiones normalizadas para matching
    medio_norm: str = field(default="", repr=False)
    formato_norm: str = field(default="", repr=False)


def _normalize(text) -> str:
    if not isinstance(text, str):
        text = "" if pd.isna(text) else str(text)
    # Collapse whitespace/newlines first
    text = " ".join(text.split())
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).upper().strip()


def _clean(val) -> str:
    if pd.isna(val):
        return ""
    return str(val).strip()


def load_base(path=None) -> list[Spec]:
    """
    Lee la hoja MATRIZ DIG de la BASE y retorna lista de Spec.
    Los headers están en la fila 2 (índice 1, 0-based).
    """
    xlsx_path = path or BASE_PATH
    df_raw = pd.read_excel(xlsx_path, sheet_name=SHEET_NAME, header=None)

    # Datos empiezan en fila 3 (índice 2)
    specs = []
    for _, row in df_raw.iloc[2:].iterrows():
        medio_val = _clean(row.iloc[COL_MEDIO] if len(row) > COL_MEDIO else "")
        formato_val = _clean(row.iloc[COL_FORMATO] if len(row) > COL_FORMATO else "")

        if not medio_val and not formato_val:
            continue

        spec = Spec(
            medio=medio_val,
            formato=formato_val,
            speccs=_clean(row.iloc[COL_SPECCS] if len(row) > COL_SPECCS else ""),
            peso=_clean(row.iloc[COL_PESO] if len(row) > COL_PESO else ""),
            extension=_clean(row.iloc[COL_EXTENSION] if len(row) > COL_EXTENSION else ""),
            texto=_clean(row.iloc[COL_TEXTO] if len(row) > COL_TEXTO else ""),
            segundaje_ideal=_clean(row.iloc[COL_SEGUNDAJE] if len(row) > COL_SEGUNDAJE else ""),
            medio_norm=_normalize(medio_val),
            formato_norm=_normalize(formato_val),
        )
        specs.append(spec)

    if not specs:
        raise ValueError(f"No se encontraron specs en la hoja '{SHEET_NAME}' de la BASE.")

    return specs
