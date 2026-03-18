"""
flow_parser.py — Parseo del FLOW y deduplicación de combinaciones únicas.
"""
import unicodedata
import pandas as pd


STOP_WORDS = {"TOTAL", "FEE", "IVA", "SUBTOTAL"}

# Columnas requeridas (normalizadas) para detectar la fila de encabezado
REQUIRED_HEADERS = {"ETAPA", "MEDIO", "FORMATO", "SEGMENTACION"}


def _normalize(text: str) -> str:
    """Elimina tildes y convierte a mayúsculas."""
    if not isinstance(text, str):
        return ""
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).upper().strip()


def _find_header_row(df_raw: pd.DataFrame) -> int:
    """
    Escanea las primeras 30 filas buscando la primera que contenga
    TODAS las columnas requeridas (normalizadas).
    Devuelve el índice de fila (0-based).
    """
    for i in range(min(30, len(df_raw))):
        row_values = {_normalize(str(v)) for v in df_raw.iloc[i] if pd.notna(v)}
        if REQUIRED_HEADERS.issubset(row_values):
            return i
    raise ValueError(
        f"No se encontró una fila de encabezado con las columnas {REQUIRED_HEADERS} "
        "en las primeras 30 filas del archivo."
    )


def _build_col_map(header_row: pd.Series) -> dict:
    """
    Mapea nombres de columna normalizados → índice de columna.
    Busca también SUB ETAPA / SUBESTAPA / SUB-ETAPA y MENSAJE.
    """
    col_map = {}
    for idx, val in enumerate(header_row):
        norm = _normalize(str(val))
        col_map[norm] = idx

    # Aliases posibles para SUB ETAPA
    for alias in ("SUB ETAPA", "SUBESTAPA", "SUB-ETAPA", "SUBETAPA"):
        if alias in col_map:
            col_map["SUB ETAPA"] = col_map[alias]
            break

    # Aliases posibles para MENSAJE/COMUNICACION
    # Incluye variantes con/sin tilde (normalizadas), en español e inglés
    _mensaje_aliases = (
        # Español directo
        "MENSAJE", "MENSAJES",
        # Comunicación (con y sin tilde → ambas llegan normalizadas sin tilde)
        "COMUNICACION", "COMUNICACIONES",
        # Inglés
        "MESSAGE", "MESSAGES",
        # Creativos / agencia
        "COPY", "COPIES",
        "CONCEPTO", "CONCEPTOS",
        "PIEZA", "PIEZAS",
        "TEMA", "TEMAS",
        "CLAIM",
        # Otros nombres que aparecen en FLOWs
        "CONTENIDO", "CONTENIDOS",
        "DESCRIPCION", "DESCRIPCION DEL MATERIAL",
        "MATERIAL", "NOMBRE MATERIAL",
    )
    for alias in _mensaje_aliases:
        if alias in col_map:
            col_map["MENSAJE"] = col_map[alias]
            break

    return col_map


def parse_flow(file_path_or_buffer) -> pd.DataFrame:
    """
    Lee el FLOW de medios y retorna un DataFrame deduplicado con columnas:
      etapa | sub_etapa | medio | formato

    Parameters
    ----------
    file_path_or_buffer : str o BytesIO
        Ruta al archivo .xlsx o buffer en memoria.
    """
    # Leer todas las hojas para detectar cuál contiene los headers correctos
    xl = pd.ExcelFile(file_path_or_buffer)
    df_raw = None
    last_error = None

    for sheet in xl.sheet_names:
        candidate = pd.read_excel(xl, header=None, sheet_name=sheet)
        try:
            _find_header_row(candidate)
            df_raw = candidate
            break
        except ValueError as e:
            last_error = e
            continue

    if df_raw is None:
        raise ValueError(
            f"No se encontró una hoja con columnas {REQUIRED_HEADERS} en las primeras 30 filas. "
            f"Hojas disponibles: {xl.sheet_names}. Último error: {last_error}"
        )

    header_idx = _find_header_row(df_raw)
    col_map = _build_col_map(df_raw.iloc[header_idx])

    # Columnas requeridas
    for req in ("ETAPA", "MEDIO", "FORMATO"):
        if req not in col_map:
            raise ValueError(f"Columna requerida '{req}' no encontrada en el encabezado.")

    etapa_col = col_map["ETAPA"]
    sub_etapa_col = col_map.get("SUB ETAPA")
    medio_col = col_map["MEDIO"]
    formato_col = col_map["FORMATO"]
    mensaje_col = col_map.get("MENSAJE")

    rows = []
    last_etapa = ""
    last_sub_etapa = ""
    last_medio = ""
    last_mensaje = ""

    data_rows = df_raw.iloc[header_idx + 1 :]

    for _, row in data_rows.iterrows():
        etapa_val = str(row.iloc[etapa_col]).strip() if pd.notna(row.iloc[etapa_col]) else ""
        sub_etapa_val = (
            str(row.iloc[sub_etapa_col]).strip()
            if sub_etapa_col is not None and pd.notna(row.iloc[sub_etapa_col])
            else ""
        )
        medio_val = str(row.iloc[medio_col]).strip() if pd.notna(row.iloc[medio_col]) else ""
        formato_val = str(row.iloc[formato_col]).strip() if pd.notna(row.iloc[formato_col]) else ""
        mensaje_val = (
            str(row.iloc[mensaje_col]).strip()
            if mensaje_col is not None and pd.notna(row.iloc[mensaje_col])
            else ""
        )

        # Detener si encontramos palabras de cierre
        norm_etapa = _normalize(etapa_val)
        norm_medio = _normalize(medio_val)
        if any(sw in norm_etapa or sw in norm_medio for sw in STOP_WORDS):
            break

        # Saltar filas completamente vacías
        if not etapa_val and not medio_val and not formato_val:
            continue

        # Saltar filas marcadoras de bloque (ETAPA definida pero MEDIO y FORMATO vacíos)
        # IMPORTANTE: este chequeo va ANTES de actualizar last_etapa/last_medio
        # para que filas de subtotal no contaminen la herencia
        if etapa_val and not medio_val and not formato_val:
            continue

        # Herencia hacia abajo: ETAPA
        if etapa_val:
            last_etapa = etapa_val
        else:
            etapa_val = last_etapa

        # Herencia: SUB ETAPA
        if sub_etapa_val:
            last_sub_etapa = sub_etapa_val
        else:
            sub_etapa_val = last_sub_etapa

        # Herencia: MEDIO
        if medio_val:
            last_medio = medio_val
        else:
            medio_val = last_medio

        # Herencia: MENSAJE (merged cells → solo primera fila tiene valor)
        if mensaje_val:
            last_mensaje = mensaje_val
        else:
            mensaje_val = last_mensaje

        # Saltar filas sin formato
        if not formato_val:
            continue

        # Limpiar saltos de línea en campos de texto simples
        etapa_val = " ".join(etapa_val.split())
        sub_etapa_val = " ".join(sub_etapa_val.split())
        medio_val = " ".join(medio_val.split())

        # FORMATO: split por \n para expandir celdas con múltiples formatos
        # Cada línea es un formato independiente que necesita su propia fila
        formato_lines = [line.strip() for line in formato_val.split("\n") if line.strip()]
        if not formato_lines:
            continue

        for fmt_line in formato_lines:
            # Normalizar espacios y luego quitar espacios alrededor del "/"
            # para que "VIDEO AD /REEL" == "VIDEO AD/REEL" en el matcher
            fmt_clean = " ".join(fmt_line.split())
            fmt_clean = "/".join(p.strip() for p in fmt_clean.split("/"))
            rows.append(
                {
                    "etapa": etapa_val,
                    "sub_etapa": sub_etapa_val,
                    "medio": medio_val,
                    "formato": fmt_clean,
                    "mensaje": mensaje_val,
                }
            )

    if not rows:
        raise ValueError("No se extrajeron filas del FLOW. Verificar estructura del archivo.")

    df = pd.DataFrame(rows)
    df = df.drop_duplicates(
        subset=["etapa", "sub_etapa", "medio", "formato", "mensaje"], keep="first"
    )
    df = df.reset_index(drop=True)
    return df
