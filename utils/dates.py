import re
from datetime import datetime
from typing import Optional, Tuple

from utils.text import safe_str


MONTHS_BR = [
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
]

RX_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
RX_BR_DATE = re.compile(r"^\s*(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{2,4}))?\s*$")
RX_TE = re.compile(
    r"(?i)(?<![A-Z0-9])TE\s*[-:\s\u2013\u2014]*\s*(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{2,4}))?"
)
RX_TE_DATE_FLEX = re.compile(
    r"\bTE\b\s*[-:\u2013\u2014]?\s*(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{2,4}))?\b",
    re.IGNORECASE,
)


def data_extenso(dt):
    """Retorna a data por extenso em portugues."""
    return f"{dt.day} de {MONTHS_BR[dt.month - 1]} de {dt.year}"


def parse_user_date(date_str: str) -> Optional[datetime]:
    text = safe_str(date_str)
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt)
        except Exception:
            pass
    return None


def parse_period_date(date_str: str, label: str) -> datetime:
    """
    Mantem compatibilidade com input type=date (YYYY-MM-DD) e aceita dd/mm/aa|aaaa.
    """
    text = safe_str(date_str)
    if not text:
        raise ValueError(f"Informe {label}.")
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt)
        except Exception:
            pass
    raise ValueError(f"Formato de {label} inválido: '{date_str}'.")


def parse_date_flexible(value: str, *, default_year: Optional[int] = None, field_label: str = "data") -> datetime:
    """
    Aceita YYYY-MM-DD, dd/mm/aaaa, dd/mm/aa ou dd/mm.
    """
    if value is None or str(value).strip() == "":
        raise ValueError(f"Informe {field_label}.")

    text = str(value).strip()

    if RX_ISO_DATE.match(text):
        try:
            return datetime.strptime(text, "%Y-%m-%d")
        except Exception:
            raise ValueError(f"{field_label.capitalize()} inválida: '{value}'.")

    match = RX_BR_DATE.match(text)
    if not match:
        raise ValueError(
            f"{field_label.capitalize()} inválida: '{value}'. "
            "Use 16/01, 16/01/26, 16/01/2026 (ou selecione no calendário)."
        )

    day = int(match.group(1))
    month = int(match.group(2))
    year_text = match.group(3)

    if not year_text:
        year = int(default_year) if default_year is not None else datetime.now().year
    else:
        year = int(year_text)
        if year < 100:
            year += 2000

    try:
        return datetime(year, month, day)
    except ValueError:
        raise ValueError(f"{field_label.capitalize()} inválida: '{value}' (dia/mês não existe).")


def extract_te_date_from_text(text: str, period_start: datetime, period_end: datetime):
    """
    Extrai TE - dd/mm[/aa|aaaa] de um texto.
    Ano ausente: tenta encaixar no periodo; se nao der, assume ano corrente.
    Retorna (dt, match_txt, year_inferred).
    """
    value = safe_str(text)
    if not value:
        return None, None, False

    match = RX_TE.search(value)
    if not match:
        return None, None, False

    day = int(match.group(1))
    month = int(match.group(2))
    year_text = match.group(3)

    year_inferred = False
    if year_text:
        year = int(year_text)
        if year < 100:
            year += 2000
        years_to_try = [year]
    else:
        year_inferred = True
        years_to_try = [period_start.year]
        if period_end.year != period_start.year:
            years_to_try.append(period_end.year)

    for year in years_to_try:
        try:
            dt = datetime(year, month, day)
        except Exception:
            continue
        if period_start <= dt <= period_end:
            return dt, match.group(0), year_inferred

    if not year_text:
        year = datetime.now().year
        try:
            return datetime(year, month, day), match.group(0), year_inferred
        except Exception:
            return None, match.group(0), year_inferred

    for year in years_to_try:
        try:
            return datetime(year, month, day), match.group(0), year_inferred
        except Exception:
            continue

    return None, match.group(0), year_inferred


def detect_te_date_from_obs_flexible(
    obs_text,
    *,
    default_year: Optional[int] = None,
) -> Tuple[Optional[datetime], Optional[str], Optional[str], bool]:
    """
    Procura TE + data em OBS.
    Retorna: (dt, regra, trecho_match, year_inferred).
    """
    if obs_text is None:
        return None, None, None, False

    text = str(obs_text).strip()
    if text == "":
        return None, None, None, False

    match = RX_TE_DATE_FLEX.search(text)
    if not match:
        return None, None, None, False

    day = int(match.group(1))
    month = int(match.group(2))
    year_text = match.group(3)

    year_inferred = False
    if not year_text:
        year = int(default_year) if default_year is not None else datetime.now().year
        year_inferred = True
    else:
        year = int(year_text)
        if year < 100:
            year += 2000

    try:
        dt = datetime(year, month, day)
    except ValueError:
        return None, "OBS:TE_DATE", match.group(0), year_inferred

    return dt, "OBS:TE_DATE", match.group(0), year_inferred
