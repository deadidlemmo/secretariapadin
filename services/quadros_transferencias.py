import re

from openpyxl.utils import get_column_letter

from utils.excel import set_merged_cell_value
from utils.text import is_missing_value, norm_header_compact, safe_str


RX_EJA_TRANSFER = re.compile(
    r"(?i)(?<![A-Z0-9])(TE|MC|MCC)\s*[-:\s\u2013\u2014]*\s*(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{2,4}))?"
)


def label_set(ws, addr: str, label: str, value: str):
    """
    Mantem prefixo do template quando existir, por exemplo:
    'Unidade Escolar: ...'.
    """
    current = ws[addr].value
    value_text = safe_str(value)

    if isinstance(current, str) and ":" in current:
        left = current.split(":", 1)[0].strip()
        if norm_header_compact(left) == norm_header_compact(label):
            set_merged_cell_value(ws, addr, f"{left}: {value_text}")
            return

    set_merged_cell_value(ws, addr, value_text)


def push_missing_info_alert(alert_items, seen_keys, *, turma, nome, ra, tipo, data_str, campo, detalhe):
    """
    Guarda inconsistencias sem duplicar o mesmo caso.
    """
    key = (
        safe_str(turma),
        safe_str(nome),
        safe_str(ra),
        safe_str(tipo),
        safe_str(data_str),
        safe_str(campo),
        safe_str(detalhe),
    )
    if key in seen_keys:
        return
    seen_keys.add(key)

    alert_items.append(
        {
            "turma": safe_str(turma) or "-",
            "nome": safe_str(nome) or "-",
            "ra": safe_str(ra) or "-",
            "tipo": safe_str(tipo) or "-",
            "data": safe_str(data_str) or "-",
            "campo": safe_str(campo) or "-",
            "detalhe": safe_str(detalhe) or "-",
        }
    )


def replace_workbook_sheet(wb, title: str):
    if title in wb.sheetnames:
        wb.remove(wb[title])
    return wb.create_sheet(title)


def add_transfer_alerts_sheet(wb, alert_items):
    if not alert_items:
        return
    ws_alert = replace_workbook_sheet(wb, "ALERTAS")
    ws_alert.append(["Turma", "Nome", "RA", "Tipo", "Data", "Campo", "Detalhe"])
    for item in alert_items:
        ws_alert.append(
            [
                item.get("turma", "-"),
                item.get("nome", "-"),
                item.get("ra", "-"),
                item.get("tipo", "-"),
                item.get("data", "-"),
                item.get("campo", "-"),
                item.get("detalhe", "-"),
            ]
        )
    for col in range(1, 8):
        ws_alert.column_dimensions[get_column_letter(col)].width = 24


def serie_key_from_value(serie_val: str):
    """Extrai 2o/3o/4o/5o de valores como '4oF', '5o D', etc."""
    text = "" if serie_val is None else str(serie_val).strip()
    match = re.search("(?<!\\d)([2345])\\s*[\\u00ba\\u00b0o]", text, flags=re.IGNORECASE)
    if not match:
        return None
    number = match.group(1)
    if number in {"2", "3", "4", "5"}:
        return f"{number}\u00ba"
    return None


def normalize_tipo_te(value) -> str:
    """
    Normaliza o TIPO TE para bater no mapping do modelo.
    Fora disso, retorna o valor original.
    """
    if is_missing_value(value):
        return "Sem Informação"

    raw = str(value).strip()
    normalized = norm_header_compact(raw)

    if "DENTRO" in normalized or "REDEMUNICIPAL" in normalized or "MUNICIPAL" in normalized:
        return "Dentro da Rede"
    if "ESTAD" in normalized:
        return "Rede Estadual"
    if "LITORAL" in normalized or "BAIXADA" in normalized:
        return "Litoral"
    if "MUDANCA" in normalized and "MUNICIP" in normalized:
        return "Mudança de Municipio"
    if "SAOPAULO" in normalized:
        return "São Paulo"
    if "ABCD" in normalized:
        return "ABCD"
    if "INTERIOR" in normalized:
        return "Interior"
    if "OUTROSESTADOS" in normalized or ("OUTROS" in normalized and "ESTAD" in normalized):
        return "Outros Estados"
    if "PARTICULAR" in normalized:
        return "Particular"
    if "PAIS" in normalized:
        return "País"
    if "SEMINFORMA" in normalized:
        return "Sem Informação"

    return raw
