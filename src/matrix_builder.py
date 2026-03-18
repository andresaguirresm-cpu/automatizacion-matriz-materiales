"""
matrix_builder.py — Construye la lista de OutputRow a partir del FLOW parseado y los specs.
"""
from dataclasses import dataclass, field
from typing import Optional
import pandas as pd

from .base_loader import Spec, load_base
from .matcher import match_formato


@dataclass
class OutputRow:
    etapa: str
    sub_etapa: str
    medio: str
    formato: str
    mensaje: str = ""
    speccs: str = ""
    peso: str = ""
    extension: str = ""
    texto: str = ""
    segundaje_ideal: str = ""


def build_matrix(
    flow_df: pd.DataFrame,
    specs: list[Spec],
) -> tuple[list[OutputRow], list[str]]:
    """
    Cruza el FLOW con la BASE y construye la lista de OutputRow.

    Parameters
    ----------
    flow_df : DataFrame con columnas [etapa, sub_etapa, medio, formato]
    specs   : lista de Spec cargados desde la BASE

    Returns
    -------
    (rows, all_warnings)
    """
    all_warnings: list[str] = []
    rows: list[OutputRow] = []

    for _, flow_row in flow_df.iterrows():
        etapa = flow_row["etapa"]
        sub_etapa = flow_row["sub_etapa"]
        medio = flow_row["medio"]
        formato = flow_row["formato"]
        mensaje = flow_row.get("mensaje", "") if hasattr(flow_row, "get") else flow_row["mensaje"]

        found_specs, warnings = match_formato(medio, formato, specs)
        all_warnings.extend(warnings)

        for i, spec in enumerate(found_specs):
            # Si hay slash-expansion, el formato de la fila es el sub-formato
            if len(found_specs) > 1:
                sub_formatos = [p.strip() for p in formato.split("/") if p.strip()]
                row_formato = sub_formatos[i] if i < len(sub_formatos) else formato
            else:
                row_formato = formato

            row = OutputRow(
                etapa=etapa,
                sub_etapa=sub_etapa,
                medio=medio,
                formato=row_formato,
                mensaje=mensaje,
                speccs=spec.speccs if spec else "",
                peso=spec.peso if spec else "",
                extension=spec.extension if spec else "",
                texto=spec.texto if spec else "",
                segundaje_ideal=spec.segundaje_ideal if spec else "",
            )
            rows.append(row)

    return rows, all_warnings
