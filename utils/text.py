import re
import unicodedata

import pandas as pd


def safe_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def norm_header_compact(text: str) -> str:
    """
    Normaliza cabecalho: remove acentos e tudo que nao for A-Z0-9.
    Fica robusto para variacoes como 'LOCAL TE', 'LOCAL_TE' e 'local te'.
    """
    if text is None:
        return ""
    normalized = str(text).strip()
    normalized = unicodedata.normalize("NFKD", normalized)
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    normalized = normalized.upper()
    return re.sub(r"[^A-Z0-9]+", "", normalized)


def build_colmap(df) -> dict:
    """Mapeia cabecalho normalizado para o nome real da coluna."""
    mapping = {}
    for col in df.columns:
        key = norm_header_compact(col)
        if key and key not in mapping:
            mapping[key] = col
    return mapping


def pick_col(colmap: dict, *candidates: str):
    """Retorna o nome real da coluna a partir de candidatos normalizados."""
    for candidate in candidates:
        key = norm_header_compact(candidate)
        if key in colmap:
            return colmap[key]
    return None


def find_df_col(df, candidates):
    if df is None or df.empty:
        return None
    return pick_col(build_colmap(df), *list(candidates))


def is_missing_value(value) -> bool:
    text = safe_str(value)
    if not text:
        return True
    return text.lower() in {"0", "-", "nan", "none", "null"}


def is_missing_text(value) -> bool:
    return is_missing_value(value)
