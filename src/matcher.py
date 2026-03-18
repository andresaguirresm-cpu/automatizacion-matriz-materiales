"""
matcher.py — Cruce FLOW → BASE con alias map, slash-expansion y fallback fuzzy.
"""
import unicodedata
from typing import Optional
from rapidfuzz import process, fuzz

from .base_loader import Spec


FUZZY_MIN_SCORE = 70

# Formatos cuyos specs son idénticos en todos los medios.
# Si (medio, formato) no se encuentra en la BASE, se busca el spec
# usando solo el formato (ignorando el medio).
CROSS_PLATFORM_FORMATS: set[str] = {
    "INSTREAM",
}

# Mapa de alias: (MEDIO_norm, FORMATO_norm) → FORMATO_norm en la BASE
ALIAS_MAP: dict[tuple[str, str], str] = {
    ("META", "PROMOTED VIDEO"): "VIDEO AD",
    ("META", "REEL"): "STORIES/REELS",
    ("META", "REELS"): "STORIES/REELS",
    ("META", "ESTATICO"): "POST AD",
    ("META", "ESTATICOS"): "POST AD",
    ("META", "STORIES"): "STORIES/REELS",
    ("META", "STORY"): "STORIES/REELS",
    ("META", "VIDEO AD"): "VIDEO AD",
    ("META", "DINAMICOS"): "CREATIVOS DINAMICOS",
    ("META", "DINAMICO"): "CREATIVOS DINAMICOS",
    ("META", "CREATIVOS DINAMICOS"): "CREATIVOS DINAMICOS",
    ("TIKTOK", "SPARK AD"): "SPARK AD (ORGANICO)",
    ("TIKTOK", "TOPVIEW"): "TOPVIEW",
    ("TIKTOK", "INFEED"): "INFEED",
    ("DV360", "INSTREAM"): "INSTREAM",
    ("DV360", "INSTREAM CONTEXTUAL"): "INSTREAM",
    ("DV360", "CONTEXTUAL"): "DISPLAY",
    ("DV360", "NATIVE ADS"): "DISPLAY",
    ("DV360", "NATIV"): "DISPLAY",
    ("DV360", "INSTREAM CONTEXTUAL NATIVE ADS"): "INSTREAM",
    ("DV360", "BANNER"): "DISPLAY",
    ("YOUTUBE", "INSTREAM"): "INSTREAM",
    ("YOUTUBE", "SHORTS"): "SHORTS",
    ("YOUTUBE", "BUMPER"): "BUMPER AD",
    ("YOUTUBE", "BUMPER AD"): "BUMPER AD",
    ("GOOGLE", "DISPLAY"): "DISPLAY",
    ("GOOGLE", "SEARCH"): "SEARCH",
    ("GOOGLE", "DEMAND GEN"): "DEMAND GEN",
}


def _normalize(text: str) -> str:
    if not isinstance(text, str):
        return ""
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).upper().strip()


def _expand_slash(formato: str) -> list[str]:
    """
    'REEL/VIDEO AD' o 'VIDEO AD /REEL/ ESTÁTICO' → ['REEL', 'VIDEO AD', 'ESTÁTICO']
    """
    parts = [p.strip() for p in formato.split("/") if p.strip()]
    return parts if len(parts) > 1 else [formato]


def _lookup_single(medio_norm: str, formato_norm: str, specs: list[Spec]) -> Optional[Spec]:
    """Busca un spec exacto por medio+formato normalizados."""
    for spec in specs:
        if spec.medio_norm == medio_norm and spec.formato_norm == formato_norm:
            return spec
    return None


def _lookup_cross_platform(formato_norm: str, specs: list[Spec]) -> Optional[Spec]:
    """
    Busca un spec por formato únicamente, ignorando el medio.
    Se usa para formatos cuyos specs son iguales en todas las plataformas (ej. INSTREAM).
    Retorna el primer match exacto encontrado.
    """
    for spec in specs:
        if spec.formato_norm == formato_norm:
            return spec
    return None


def _fuzzy_lookup(medio_norm: str, formato_norm: str, specs: list[Spec]) -> tuple[Optional[Spec], float]:
    """
    Primero filtra candidatos del mismo MEDIO, luego hace fuzzy sobre FORMATO.
    """
    candidates = [s for s in specs if s.medio_norm == medio_norm]
    if not candidates:
        # Si no hay candidatos del mismo medio, buscar en todos
        candidates = specs

    choices = {s.formato_norm: s for s in candidates}
    result = process.extractOne(
        formato_norm,
        list(choices.keys()),
        scorer=fuzz.token_set_ratio,
    )
    if result and result[1] >= FUZZY_MIN_SCORE:
        return choices[result[0]], result[1]
    return None, 0.0


def match_formato(
    medio: str,
    formato: str,
    specs: list[Spec],
) -> tuple[list[Optional[Spec]], list[str]]:
    """
    Dado un MEDIO y FORMATO del FLOW, retorna:
      - Lista de Spec encontrados (uno por sub-formato expandido)
      - Lista de warnings

    Maneja:
      1. Slash-notation expansion
      2. Alias map
      3. Búsqueda exacta
      4. Fuzzy fallback
    """
    medio_norm = _normalize(medio)
    sub_formatos = _expand_slash(formato)

    found_specs: list[Optional[Spec]] = []
    warnings: list[str] = []

    for sub_fmt in sub_formatos:
        fmt_norm = _normalize(sub_fmt)

        # 1. Alias map
        alias_fmt = ALIAS_MAP.get((medio_norm, fmt_norm))
        if alias_fmt:
            fmt_norm_lookup = _normalize(alias_fmt)
        else:
            fmt_norm_lookup = fmt_norm

        # 2. Búsqueda exacta
        spec = _lookup_single(medio_norm, fmt_norm_lookup, specs)

        # 3. Fallback cross-platform: formatos con specs compartidos entre medios
        if spec is None and fmt_norm_lookup in CROSS_PLATFORM_FORMATS:
            spec = _lookup_cross_platform(fmt_norm_lookup, specs)
            if spec:
                warnings.append(
                    f"Formato '{sub_fmt}' (Medio: {medio}) → specs tomados de "
                    f"'{spec.medio}' (formato cross-platform)."
                )

        # 4. Fuzzy fallback
        if spec is None:
            spec, score = _fuzzy_lookup(medio_norm, fmt_norm_lookup, specs)
            if spec:
                warnings.append(
                    f"Formato '{sub_fmt}' (Medio: {medio}) → match fuzzy con "
                    f"'{spec.formato}' (score={score:.0f}). Verificar."
                )
            else:
                warnings.append(
                    f"Formato '{sub_fmt}' (Medio: {medio}) no encontrado en la BASE. "
                    "Las columnas de specs quedarán vacías."
                )

        found_specs.append(spec)

    return found_specs, warnings
