import re
import string
from datetime import datetime

from utils.excel import set_merged_cell_value


TURMAS_MAX = list(string.ascii_uppercase[:14])  # A..N

ATENDIMENTO_CONFIG = {
    "MODEL_BLOCKS": {
        "1\u00ba": {"start_row": 19, "masc_col": "B", "fem_col": "C", "total_col": "D"},
        "2\u00ba": {"start_row": 37, "masc_col": "B", "fem_col": "C", "total_col": "D"},
        "3\u00ba": {"start_row": 55, "masc_col": "B", "fem_col": "C", "total_col": "D"},
        "4\u00ba": {"start_row": 73, "masc_col": "B", "fem_col": "C", "total_col": "D"},
        "5\u00ba": {"start_row": 91, "masc_col": "B", "fem_col": "C", "total_col": "D"},
    },
    "PILOTO_COLS_DEFAULT": {
        "serie_col": 3,
        "turma_col": 4,
        "ma_masc_col": 7,
        "ma_fem_col": 8,
        "ma_total_col": 9,
    },
    "FALLBACK_START": {
        "2\u00ba": {"row": 6, "masc_col": 7, "fem_col": 8},
        "3\u00ba": {"row": 14, "masc_col": 7, "fem_col": 8},
        "4\u00ba": {"row": 21, "masc_col": 7, "fem_col": 8},
        "5\u00ba": {"row": 29, "masc_col": 7, "fem_col": 8},
    },
    "TOTALS": {
        "manha_total_cell": (38, 9),
        "tarde_total_cell": (40, 9),
        "modelo_manha_addr": "R20",
        "modelo_tarde_addr": "R28",
    },
    "ENABLE_DEBUG_LOG": True,
}

EJA_ZERO_CELLS = [
    "L19", "L20", "L21", "L22",
    "M19", "M20", "M21", "M22",
    "L27", "L28", "L29", "L30",
    "M27", "M28", "M29", "M30",
    "L35", "L36", "L37",
    "M35", "M36", "M37",
    "R32",
]


def normalize_mes_ref(value, now=None) -> str:
    now = now or datetime.now()
    text = (value or "").strip()
    if text:
        match = re.match(r"^\s*(\d{4})-(\d{2})\s*$", text)
        if match:
            return f"{match.group(2)}/{match.group(1)}"

        match = re.match(r"^\s*(\d{2})/(\d{4})\s*$", text)
        if match:
            return f"{match.group(1)}/{match.group(2)}"

    return now.strftime("%m/%Y")


def safe_int(value, default=0):
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        try:
            return int(value)
        except Exception:
            return default
    try:
        text = str(value).strip()
        text = text.replace(".", "").replace(",", ".")
        return int(float(text))
    except Exception:
        return default


def norm_serie(value):
    if value is None:
        return None
    text = str(value).strip()
    match = re.search(r"(\d)\s*[\u00bao\u00b0]?", text, flags=re.IGNORECASE)
    if not match:
        return None
    number = match.group(1)
    if number in {"1", "2", "3", "4", "5"}:
        return f"{number}\u00ba"
    return None


def extract_turma_letter(value):
    if value is None:
        return None
    text = str(value).strip().upper()
    match = re.search(r"\b([A-N])\b", text)
    if match:
        return match.group(1)
    fallback = re.search(r"([A-N])", text)
    return fallback.group(1) if fallback else None


def condense_letters(letters):
    if not letters:
        return "-"

    indexes = sorted({TURMAS_MAX.index(letter) for letter in letters if letter in TURMAS_MAX})
    output = []
    i = 0
    while i < len(indexes):
        start = indexes[i]
        j = i
        while j + 1 < len(indexes) and indexes[j + 1] == indexes[j] + 1:
            j += 1
        if j == i:
            output.append(TURMAS_MAX[start])
        else:
            output.append(f"{TURMAS_MAX[start]}-{TURMAS_MAX[indexes[j]]}")
        i = j + 1

    return ", ".join(output)


def detect_ma_columns(ws_total):
    cfg = ATENDIMENTO_CONFIG["PILOTO_COLS_DEFAULT"].copy()
    try:
        for row in range(1, 20):
            for col in range(1, 30):
                value = ws_total.cell(row=row, column=col).value
                if value and "MATRICULAS" in str(value).upper():
                    cfg["ma_masc_col"] = col
                    cfg["ma_fem_col"] = col + 1
                    cfg["ma_total_col"] = col + 2
                    return cfg
    except Exception:
        pass
    return cfg


def extract_by_cols(ws_total, serie_label, debug_log):
    cols = detect_ma_columns(ws_total)
    serie_col = cols["serie_col"]
    turma_col = cols["turma_col"]
    masc_col = cols["ma_masc_col"]
    fem_col = cols["ma_fem_col"]

    found = {}
    duplicates = []
    limit = min(ws_total.max_row or 0, 300)

    for row in range(1, limit + 1):
        serie = norm_serie(ws_total.cell(row=row, column=serie_col).value)
        if serie != serie_label:
            continue

        turma = extract_turma_letter(ws_total.cell(row=row, column=turma_col).value)
        if not turma:
            continue

        masc = safe_int(ws_total.cell(row=row, column=masc_col).value, 0)
        fem = safe_int(ws_total.cell(row=row, column=fem_col).value, 0)

        if turma in found:
            duplicates.append(turma)
            found[turma] = (found[turma][0] + masc, found[turma][1] + fem)
        else:
            found[turma] = (masc, fem)

    if duplicates:
        debug_log.append(f"[{serie_label}] AVISO: turmas duplicadas (somadas): {sorted(set(duplicates))}")

    return found


def extract_by_fallback_block(ws_total, serie_label, debug_log):
    fallback = ATENDIMENTO_CONFIG["FALLBACK_START"].get(serie_label)
    if not fallback:
        return {}

    serie_col = ATENDIMENTO_CONFIG["PILOTO_COLS_DEFAULT"]["serie_col"]
    turma_col = ATENDIMENTO_CONFIG["PILOTO_COLS_DEFAULT"]["turma_col"]
    row = fallback["row"]
    masc_col = fallback["masc_col"]
    fem_col = fallback["fem_col"]

    found = {}
    for _ in range(len(TURMAS_MAX)):
        serie_here = norm_serie(ws_total.cell(row=row, column=serie_col).value)
        if serie_here != serie_label:
            break

        turma = extract_turma_letter(ws_total.cell(row=row, column=turma_col).value)
        if not turma:
            break

        masc = safe_int(ws_total.cell(row=row, column=masc_col).value, 0)
        fem = safe_int(ws_total.cell(row=row, column=fem_col).value, 0)
        found[turma] = (masc, fem)
        row += 1

    debug_log.append(
        f"[{serie_label}] fallback usado: capturadas {len(found)} turmas "
        "(parada por mudan\u00e7a de s\u00e9rie na col C)."
    )
    return found


def write_block(ws_modelo, serie_label, turma_data, debug_log):
    block = ATENDIMENTO_CONFIG["MODEL_BLOCKS"][serie_label]
    start = block["start_row"]
    masc_col = block["masc_col"]
    fem_col = block["fem_col"]
    total_col = block["total_col"]

    found_letters = [turma for turma in TURMAS_MAX if turma in turma_data]
    missing_letters = [turma for turma in TURMAS_MAX if turma not in turma_data]

    for index, turma in enumerate(TURMAS_MAX):
        row = start + index
        masc = turma_data.get(turma, (0, 0))[0]
        fem = turma_data.get(turma, (0, 0))[1]

        set_merged_cell_value(ws_modelo, f"{masc_col}{row}", masc)

        if fem_col:
            set_merged_cell_value(ws_modelo, f"{fem_col}{row}", fem)

        if total_col and fem_col:
            set_merged_cell_value(ws_modelo, f"{total_col}{row}", f"={masc_col}{row}+{fem_col}{row}")

    debug_log.append(
        f"[{serie_label}] preenchido: {condense_letters(found_letters)}; "
        f"zerado: {condense_letters(missing_letters)}"
    )


def read_total(ws_total, row, col, debug_log, label):
    value = ws_total.cell(row=row, column=col).value
    output = safe_int(value, 0)
    debug_log.append(f"[TOTAL] {label}: lido {output} de ({row},{col}).")
    return output


def write_header(ws_modelo, responsavel, rf, mes_ref, debug_log):
    set_merged_cell_value(ws_modelo, "B5", "E.M Jos\u00e9 Padin Mouta")
    set_merged_cell_value(ws_modelo, "C6", responsavel or "-")
    set_merged_cell_value(ws_modelo, "B7", rf or "-")
    set_merged_cell_value(ws_modelo, "A13", mes_ref)
    debug_log.append(f"[HEADER] responsavel='{responsavel or '-'}' rf='{rf or '-'}' mes_ref='{mes_ref}'")


def write_turno_totals(ws_modelo, ws_total, debug_log):
    config = ATENDIMENTO_CONFIG["TOTALS"]
    manha = read_total(ws_total, *config["manha_total_cell"], debug_log=debug_log, label="Manh\u00e3")
    tarde = read_total(ws_total, *config["tarde_total_cell"], debug_log=debug_log, label="Tarde")

    set_merged_cell_value(ws_modelo, config["modelo_manha_addr"], manha)
    set_merged_cell_value(ws_modelo, "R24", "-")
    set_merged_cell_value(ws_modelo, config["modelo_tarde_addr"], tarde)

    return manha, tarde


def zero_eja_block(ws_modelo):
    for address in EJA_ZERO_CELLS:
        set_merged_cell_value(ws_modelo, address, 0)
    set_merged_cell_value(ws_modelo, "R24", "-")


def fill_eja_block(ws_modelo, ws_total_eja):
    for source_row, target_row in ((6, 19), (7, 20), (8, 21), (9, 22)):
        set_merged_cell_value(ws_modelo, f"L{target_row}", ws_total_eja.cell(row=source_row, column=5).value)
        set_merged_cell_value(ws_modelo, f"M{target_row}", ws_total_eja.cell(row=source_row, column=6).value)

    for source_row, target_row in ((11, 27), (12, 28), (13, 29), (14, 30)):
        set_merged_cell_value(ws_modelo, f"L{target_row}", ws_total_eja.cell(row=source_row, column=5).value)
        set_merged_cell_value(ws_modelo, f"M{target_row}", ws_total_eja.cell(row=source_row, column=6).value)

    for source_row, target_row in ((16, 35), (17, 36), (18, 37)):
        set_merged_cell_value(ws_modelo, f"L{target_row}", ws_total_eja.cell(row=source_row, column=5).value)
        set_merged_cell_value(ws_modelo, f"M{target_row}", ws_total_eja.cell(row=source_row, column=6).value)

    set_merged_cell_value(ws_modelo, "R32", ws_total_eja.cell(row=20, column=7).value)
    set_merged_cell_value(ws_modelo, "R24", "-")
