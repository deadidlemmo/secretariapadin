from datetime import datetime

import pandas as pd

from services.carteirinhas_log import get_printed_set, mark_printed_rms
from services.fotos import get_student_photo_url, student_has_photo


CARTEIRINHAS_COLUMNS = [
    "RM",
    "NOME",
    "DATA NASC.",
    "RA",
    "SAI SOZINHO?",
    "S\u00c9RIE",
    "HOR\u00c1RIO",
]


def normalize_rms(rms):
    normalized = []
    for value in rms or []:
        try:
            rm = int(str(value).strip())
            if rm > 0:
                normalized.append(rm)
        except Exception:
            pass
    return list(dict.fromkeys(normalized))


def _format_date_br(value) -> str:
    if pd.notna(value):
        try:
            date_value = pd.to_datetime(value, errors="coerce")
            if pd.notna(date_value):
                return date_value.strftime("%d/%m/%Y")
        except Exception:
            pass
    return "Desconhecida"


def _status_sai_sozinho(value):
    raw = str(value).strip().upper()
    if raw in ("SIM", "S", "YES", "Y"):
        return {
            "classe_cor": "verde",
            "status_texto": "Sai sozinho",
            "status_icon": "&#10003;",
        }
    return {
        "classe_cor": "vermelho",
        "status_texto": "N\u00e3o sai sozinho",
        "status_icon": "&#9888;",
    }


def _read_carteirinhas_lista(arquivo_excel):
    planilha = pd.read_excel(arquivo_excel, sheet_name="LISTA CORRIDA").copy()
    dados = planilha.loc[:, CARTEIRINHAS_COLUMNS].copy()
    dados.loc[:, "RM"] = pd.to_numeric(dados["RM"], errors="coerce").fillna(0).astype(int)
    return dados


def _paginate(items, per_page=6):
    return [items[index:index + per_page] for index in range(0, len(items), per_page)]


def build_carteirinhas_context(
    arquivo_excel,
    somente_com_foto=False,
    somente_nao_impressas=False,
    ano=None,
    get_printed_set_func=get_printed_set,
    get_photo_url_func=get_student_photo_url,
):
    ano = int(ano or datetime.now().year)
    printed_set = get_printed_set_func(ano)
    dados = _read_carteirinhas_lista(arquivo_excel)

    alunos_sem_fotos_list = []
    alunos = []

    for _, row in dados.iterrows():
        rm_int = int(row["RM"])
        if rm_int <= 0:
            continue

        foto_url = get_photo_url_func(rm_int)
        if not foto_url:
            alunos_sem_fotos_list.append(
                {
                    "rm": rm_int,
                    "nome": row["NOME"],
                    "serie": row["S\u00c9RIE"],
                }
            )

        status = _status_sai_sozinho(row["SAI SOZINHO?"])

        alunos.append(
            {
                "rm": rm_int,
                "nome": row["NOME"],
                "data_nasc": _format_date_br(row["DATA NASC."]),
                "ra": row["RA"],
                "serie": row["S\u00c9RIE"],
                "horario": row["HOR\u00c1RIO"],
                "classe_cor": status["classe_cor"],
                "status_texto": status["status_texto"],
                "status_icon": status["status_icon"],
                "foto_url": foto_url,
                "impresso": rm_int in printed_set,
            }
        )

    alunos_para_exibir = alunos
    if somente_com_foto:
        alunos_para_exibir = [aluno for aluno in alunos_para_exibir if aluno.get("foto_url")]
    if somente_nao_impressas:
        alunos_para_exibir = [aluno for aluno in alunos_para_exibir if not aluno.get("impresso")]

    return {
        "pages": _paginate(alunos_para_exibir, per_page=6),
        "alunos_sem_foto": alunos_sem_fotos_list,
        "total_sem_foto": len(alunos_sem_fotos_list),
        "somente_com_foto": somente_com_foto,
        "somente_nao_impressas": somente_nao_impressas,
        "ano": ano,
    }


def mark_carteirinhas_impressas(
    rms,
    ano=None,
    student_has_photo_func=student_has_photo,
    mark_printed_rms_func=mark_printed_rms,
):
    ano = int(ano or datetime.now().year)
    normalized = normalize_rms(rms)
    rms_with_photo = [rm for rm in normalized if student_has_photo_func(rm)]
    added, total_printed = mark_printed_rms_func(ano, rms_with_photo)

    return {
        "ok": True,
        "ano": ano,
        "received": len(normalized),
        "considered_with_photo": len(rms_with_photo),
        "added": added,
        "total_printed": total_printed,
    }
