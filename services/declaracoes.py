import re
from datetime import datetime
from html import escape
from numbers import Number

import pandas as pd


MONTHS_BR = {
    1: "janeiro",
    2: "fevereiro",
    3: "mar\u00e7o",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}

DECLARACAO_PRINT_CSS = """
.acoes-preview {
  position: fixed;
  top: 16px;
  right: 16px;
  z-index: 9999;
  display: flex;
  gap: 8px;
}
.print-button {
  display: inline-block;
  background-color: #283E51;
  color: #fff;
  border: none;
  padding: 10px 20px;
  border-radius: 5px;
  cursor: pointer;
  text-decoration: none;
  font-family: Arial, sans-serif;
  font-size: 14px;
  line-height: 1.2;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
}
.print-button:hover {
  background-color: #1d2d3a;
  color: #fff;
  text-decoration: none;
}
.print-button-disabled {
  cursor: not-allowed;
  opacity: 0.65;
}
@media print {
  .acoes-preview {
    display: none !important;
  }
}
"""

NOTAS_MATERIAS = [
    ("L\u00edngua Portuguesa", "LP_1T", "LP_2T", "LP_3T"),
    ("Hist\u00f3ria", "HIST_1T", "HIST_2T", "HIST_3T"),
    ("Geografia", "GEO_1T", "GEO_2T", "GEO_3T"),
    ("Matem\u00e1tica", "MAT_1T", "MAT_2T", "MAT_3T"),
    ("Ci\u00eancias", "CIEN_1T", "CIEN_2T", "CIEN_3T"),
    ("Educa\u00e7\u00e3o F\u00edsica", "EDFIS_1T", "EDFIS_2T", "EDFIS_3T"),
    ("Arte", "ARTE_1T", "ARTE_2T", "ARTE_3T"),
]


def format_rm(value) -> str:
    try:
        return str(int(float(value)))
    except Exception:
        return str(value)


def format_eja_rm(value) -> str:
    try:
        if pd.notna(value) and float(value) != 0:
            return str(int(value))
    except Exception:
        pass
    return ""


def format_date_br(value, default="Desconhecida") -> str:
    if pd.notna(value):
        try:
            date_value = pd.to_datetime(value, errors="coerce")
            if pd.notna(date_value):
                return date_value.strftime("%d/%m/%Y")
        except Exception:
            pass
    return default


def format_serie_ano(value) -> str:
    if isinstance(value, str):
        return re.sub(r"(\d+\u00ba)\s*([A-Za-z])", r"\1 ano \2", value)
    return value


def read_lista_corrida_fundamental(file_path):
    planilha = pd.read_excel(file_path, sheet_name="LISTA CORRIDA").copy()
    planilha.columns = [str(column).strip().upper() for column in planilha.columns]
    planilha.loc[:, "RM_str"] = planilha["RM"].apply(format_rm)
    return planilha


def _get_eja_ra(row):
    try:
        value = row.iloc[7]
        if pd.isna(value) or float(value) == 0:
            return row.iloc[8]
        return value
    except Exception:
        return row.iloc[7]


def read_lista_eja_declaracoes(file_path):
    df = pd.read_excel(file_path, sheet_name=0, header=None, skiprows=1).copy()
    df.columns = [str(column).strip().upper() for column in df.columns]
    df.loc[:, "RM_str"] = df.iloc[:, 2].apply(format_eja_rm)
    df.loc[:, "NOME"] = df.iloc[:, 3]
    df.loc[:, "NASC."] = df.iloc[:, 6]
    df.loc[:, "RA"] = df.apply(_get_eja_ra, axis=1)
    df.loc[:, "S\u00c9RIE"] = df.iloc[:, 0]
    return df


def list_declaracao_alunos(file_path, segmento):
    if segmento == "EJA":
        df = read_lista_eja_declaracoes(file_path)
        alunos_df = df.loc[df["RM_str"] != "", ["RM_str", "NOME"]].drop_duplicates()
    else:
        df = read_lista_corrida_fundamental(file_path)
        alunos_df = df.loc[df["RM_str"] != "0", ["RM_str", "NOME"]].drop_duplicates()

    return [
        {"rm": row["RM_str"], "nome": row["NOME"]}
        for _, row in alunos_df.iterrows()
    ]


def load_declaracao_aluno_context(file_path, rm, segmento, tipo):
    rm_num = format_rm(rm)

    if segmento != "EJA":
        planilha = read_lista_corrida_fundamental(file_path)
        aluno = planilha.loc[planilha["RM_str"] == rm_num]
        if aluno.empty:
            return None

        row = aluno.iloc[0]
        serie = row["S\u00c9RIE"]
        if isinstance(serie, str):
            serie = format_serie_ano(serie)

        horario = row.get("HOR\u00c1RIO", "Desconhecido")
        if pd.isna(horario) or not str(horario).strip():
            horario = "Desconhecido"
        else:
            horario = str(horario).strip()

        notas_tabela_html = ""
        if tipo == "Transferencia":
            notas_tabela_html = build_notas_tabela_html(file_path, rm_num)

        return {
            "row": row,
            "rm_num": rm_num,
            "nome": row["NOME"],
            "serie": serie,
            "data_nasc": format_date_br(row["DATA NASC."]),
            "ra": row["RA"],
            "ra_label": "RA",
            "horario": horario,
            "semestre_texto": "",
            "notas_tabela_html": notas_tabela_html,
        }

    df = read_lista_eja_declaracoes(file_path)
    aluno = df.loc[df["RM_str"] == rm_num]
    if aluno.empty:
        return None

    row = aluno.iloc[0]
    if len(row) > 29:
        semestre = row.iloc[29]
        semestre_texto = str(semestre).strip() if pd.notna(semestre) else ""
    else:
        semestre_texto = ""

    serie = row["S\u00c9RIE"]
    if isinstance(serie, str):
        serie = format_serie_ano(serie)

    original_ra = row.iloc[7]
    if pd.isna(original_ra) or (
        isinstance(original_ra, Number) and float(original_ra) == 0
    ):
        ra_label = "RG"
    else:
        ra_label = "RA"

    return {
        "row": row,
        "rm_num": rm_num,
        "nome": row["NOME"],
        "serie": serie,
        "data_nasc": format_date_br(row["NASC."]),
        "ra": row["RA"],
        "ra_label": ra_label,
        "horario": "Desconhecido",
        "semestre_texto": semestre_texto,
        "notas_tabela_html": "",
    }


def _is_quinto_ano(value) -> bool:
    serie_raw = str(value or "").strip()
    return "5\u00ba" in serie_raw or "5\u00b0" in serie_raw


def _format_quinto_ano_serie(value) -> str:
    serie_fmt = str(value or "").strip()
    try:
        return format_serie_ano(serie_fmt)
    except Exception:
        return serie_fmt


def _base_quinto_ano_record(row):
    rm_str = str(row.get("RM_str", "")).strip()
    if rm_str in ("", "0"):
        return None

    serie_raw = str(row.get("S\u00c9RIE", "")).strip()
    if not serie_raw or not _is_quinto_ano(serie_raw):
        return None

    horario = str(row.get("HOR\u00c1RIO", "")).strip()
    if not horario:
        horario = "Desconhecido"

    return {
        "nome": str(row.get("NOME", "")).strip(),
        "ra": str(row.get("RA", "")).strip(),
        "data_nasc": format_date_br(row.get("DATA NASC.")),
        "serie_fmt": _format_quinto_ano_serie(serie_raw),
        "horario": horario,
        "valor_bolsa": str(row.get("BOLSA FAMILIA", "")).strip().upper(),
    }


def build_lote_escolaridade_5ano_context(file_path, file_path2=None, now=None):
    """
    Monta os dados das declaracoes em lote de escolaridade do 5o ano.
    """
    effective_path = file_path2 if file_path2 is not None else file_path
    if not effective_path:
        raise ValueError(
            "Caminho do arquivo Excel n\u00e3o informado para o lote de escolaridade 5\u00ba ano."
        )

    planilha = read_lista_corrida_fundamental(effective_path)
    registros = []

    for _, row in planilha.iterrows():
        record = _base_quinto_ano_record(row)
        if record is None:
            continue

        texto = (
            f"Declaro, para os devidos fins, que o(a) aluno(a) "
            f"<strong><u>{record['nome']}</u></strong>, portador(a) do RA "
            f"<strong><u>{record['ra']}</u></strong>, nascido(a) em "
            f"<strong><u>{record['data_nasc']}</u></strong>, "
            f"encontra-se regularmente matriculado(a) na "
            f"E.M Jos\u00e9 Padin Mouta, cursando atualmente o(a) "
            f"<strong><u>{record['serie_fmt']}</u></strong> no hor\u00e1rio de aula: "
            f"<strong><u>{record['horario']}</u></strong>."
        )

        registros.append(
            {
                "nome": record["nome"],
                "ra": record["ra"],
                "data_nasc": record["data_nasc"],
                "serie_fmt": record["serie_fmt"],
                "horario": record["horario"],
                "texto": texto,
            }
        )

    return {
        "registros": registros,
        "data_extenso": data_extenso_praia_grande(now),
        "titulo": "Declara\u00e7\u00e3o de Escolaridade",
    }


def _next_series_text_from_serie(serie_fmt) -> str:
    series_text = "a s\u00e9rie subsequente"
    match = re.search(r"(\d+)\u00ba", str(serie_fmt))
    if match:
        try:
            next_year = int(match.group(1)) + 1
            series_text = f"{next_year}\u00ba ano"
        except Exception:
            pass
    return series_text


def _build_conclusao_5ano_text(record, bolsa_familia_src):
    declaracao_text = (
        f"Declaro, para os devidos fins, que o(a) aluno(a) "
        f"<strong><u>{record['nome']}</u></strong>, "
        f"portador(a) do RA <strong><u>{record['ra']}</u></strong>, nascido(a) em "
        f"<strong><u>{record['data_nasc']}</u></strong>, concluiu com \u00eaxito o "
        f"<strong><u>{record['serie_fmt']}</u></strong>, estando apto(a) a ingressar no "
        f"<strong><u>{record['series_text']}</u></strong>."
    )

    if record["valor_bolsa"] == "SIM":
        declaracao_text += "<br><br><strong>Observa\u00e7\u00f5es:</strong><br>"
        declaracao_text += (
            '<label class="checkbox-label" '
            'style="display: block; text-align: justify; font-size:14px;">'
        )
        declaracao_text += (
            f'<img src="{bolsa_familia_src}" '
            'alt="Bolsa Fam\u00edlia" '
            'style="width:28px; vertical-align:middle; margin-right:5px;">'
            "O aluno \u00e9 benefici\u00e1rio do Programa Bolsa Fam\u00edlia."
        )
        declaracao_text += "</label>"

    return declaracao_text


def build_lote_conclusao_5ano_context(
    file_path,
    data_extenso,
    bolsa_familia_src="/static/logos/bolsa_familia.jpg",
):
    """
    Monta os dados das declaracoes em lote de conclusao do 5o ano.
    """
    planilha = read_lista_corrida_fundamental(file_path)
    registros = []

    for _, row in planilha.iterrows():
        record = _base_quinto_ano_record(row)
        if record is None:
            continue

        record["series_text"] = _next_series_text_from_serie(record["serie_fmt"])
        record["texto"] = _build_conclusao_5ano_text(record, bolsa_familia_src)
        registros.append(
            {
                "nome": record["nome"],
                "ra": record["ra"],
                "data_nasc": record["data_nasc"],
                "serie_fmt": record["serie_fmt"],
                "series_text": record["series_text"],
                "texto": record["texto"],
            }
        )

    registros_duas_vias = []
    for reg in registros:
        reg1 = reg.copy()
        reg1["via"] = 1
        registros_duas_vias.append(reg1)

        reg2 = reg.copy()
        reg2["via"] = 2
        registros_duas_vias.append(reg2)

    return {
        "registros": registros_duas_vias,
        "data_extenso": data_extenso,
        "titulo": "Declara\u00e7\u00e3o de Conclus\u00e3o",
        "total": len(registros_duas_vias),
    }


def data_extenso_praia_grande(now=None) -> str:
    now = now or datetime.now()
    month = MONTHS_BR[now.month].capitalize()
    return f"Praia Grande, {now.day:02d} de {month} de {now.year}"


def get_str(dados, key, default="") -> str:
    return (dados.get(key) or default).strip()


def normalizar_semestre(dados, *keys) -> str:
    for key in keys:
        value = dados.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def normalizar_segmento_personalizado(dados) -> str:
    raw = dados.get("segmento") or dados.get("segmento_personalizado") or "Fundamental"
    normalized = str(raw).strip().lower()
    if normalized in ("fundamental", "fund", "ef", "ensino fundamental"):
        return "Fundamental"
    return "EJA"


def contexto_segmento(segmento: str):
    if segmento == "Fundamental":
        return "Ensino Fundamental", "do"
    return "Educa\u00e7\u00e3o de Jovens e Adultos (EJA)", "da"


def normalizar_tipo_declaracao(dados) -> str:
    raw = dados.get("tipo_declaracao") or dados.get("tipo_declaracao_personalizada")
    return (raw or "").strip().lower()


def parse_data_nascimento_personalizada(value) -> str:
    text = (value or "").strip()
    if not text:
        return "Desconhecida"

    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except Exception:
            continue

    return "Desconhecida"


MESES_FREQUENCIA_BR = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Mar\u00e7o",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

FREQUENCIA_FORM_MESES = [
    ("jan", "Janeiro"),
    ("fev", "Fevereiro"),
    ("mar", "Mar\u00e7o"),
    ("abr", "Abril"),
    ("mai", "Maio"),
    ("jun", "Junho"),
    ("jul", "Julho"),
    ("ago", "Agosto"),
    ("set", "Setembro"),
    ("out", "Outubro"),
    ("nov", "Novembro"),
    ("dez", "Dezembro"),
]


class DeclaracaoFormError(ValueError):
    def __init__(self, message, segmento=None):
        super().__init__(message)
        self.segmento = segmento


def normalizar_tipo_escolar_form(tipo) -> str:
    text = (tipo or "").strip()
    lower = text.lower()
    if lower in ("transferencia", "transfer\u00eancia"):
        return "Transferencia"
    if lower in ("conclusao", "conclus\u00e3o"):
        return "Conclus\u00e3o"
    if lower in ("frequencia", "frequ\u00eancia"):
        return "Frequencia"
    return text


def build_declaracao_personalizada_payload(form) -> dict:
    segmento_pers = form.get("segmento_personalizado")
    if segmento_pers not in ("Fundamental", "EJA"):
        raise DeclaracaoFormError(
            "Selecione o segmento (Ensino Fundamental ou EJA) na declara\u00e7\u00e3o personalizada.",
            segmento="Personalizado",
        )

    nome_aluno = (form.get("nome_aluno") or "").strip()
    data_nascimento = form.get("data_nascimento")
    ra = (form.get("ra") or "").strip()
    tipo_pers = form.get("tipo_declaracao_personalizada")

    if (
        not nome_aluno
        or not data_nascimento
        or not ra
        or tipo_pers not in ("Conclusao", "MatriculaCancelada", "NCOM")
    ):
        raise DeclaracaoFormError(
            "Preencha todos os dados obrigat\u00f3rios da declara\u00e7\u00e3o personalizada.",
            segmento="Personalizado",
        )

    dados_personalizados = {
        "segmento": segmento_pers,
        "nome_aluno": nome_aluno,
        "data_nascimento": data_nascimento,
        "ra": ra,
        "tipo_declaracao": tipo_pers,
    }

    if tipo_pers == "Conclusao":
        ano_serie_concluida = (form.get("ano_serie_concluida") or "").strip()
        ano_conclusao = (form.get("ano_conclusao") or "").strip()
        deve_hist_unidade = form.get("deve_historico_unidade")
        semestre_conclusao = (form.get("semestre_conclusao") or "").strip()

        campos_invalidos = (
            not ano_serie_concluida
            or not ano_conclusao
            or deve_hist_unidade not in ("Sim", "N\u00e3o")
        )

        if segmento_pers == "EJA" and not semestre_conclusao:
            campos_invalidos = True

        if campos_invalidos:
            raise DeclaracaoFormError(
                "Preencha todos os campos da declara\u00e7\u00e3o personalizada de conclus\u00e3o "
                "(para EJA \u00e9 obrigat\u00f3rio informar o semestre).",
                segmento="Personalizado",
            )

        dados_personalizados.update(
            {
                "ano_serie_concluida": ano_serie_concluida,
                "ano_conclusao": ano_conclusao,
                "deve_historico_unidade": (deve_hist_unidade == "Sim"),
                "semestre_conclusao": semestre_conclusao,
            }
        )

    elif tipo_pers == "MatriculaCancelada":
        ano_serie_matricula = (form.get("ano_serie_matricula") or "").strip()
        ano_matricula = (form.get("ano_matricula") or "").strip()
        semestre_matricula = (form.get("semestre_matricula") or "").strip()

        if not ano_serie_matricula or not ano_matricula or not semestre_matricula:
            raise DeclaracaoFormError(
                "Preencha todos os campos da declara\u00e7\u00e3o de matr\u00edcula cancelada.",
                segmento="Personalizado",
            )

        dados_personalizados.update(
            {
                "ano_serie_matricula": ano_serie_matricula,
                "ano_matricula": ano_matricula,
                "semestre_matricula": semestre_matricula,
            }
        )

    elif tipo_pers == "NCOM":
        ano_serie_vaga = (form.get("ano_serie_vaga") or "").strip()
        ano_referencia_ncom = (form.get("ano_referencia_ncom") or "").strip()
        semestre_referencia_ncom = (form.get("semestre_referencia_ncom") or "").strip()

        if not ano_serie_vaga or not ano_referencia_ncom:
            raise DeclaracaoFormError(
                "Preencha todos os campos obrigat\u00f3rios da declara\u00e7\u00e3o de "
                "N\u00e3o Comparecimento (NCOM).",
                segmento="Personalizado",
            )

        dados_personalizados.update(
            {
                "ano_serie_vaga": ano_serie_vaga,
                "ano_referencia_ncom": ano_referencia_ncom,
                "semestre_referencia_ncom": semestre_referencia_ncom,
            }
        )

    return dados_personalizados


def resolve_historico_fields(tipo, deve_historico_str, unidade_select="", unidade_manual=""):
    unidade_anterior = (unidade_select or "").strip() or (unidade_manual or "").strip()

    if tipo not in ("Transferencia", "Conclus\u00e3o"):
        return False, ""

    if deve_historico_str not in ("sim", "nao"):
        raise DeclaracaoFormError("Por favor, responda se o aluno deve o hist\u00f3rico escolar.")

    if deve_historico_str == "sim" and not unidade_anterior:
        raise DeclaracaoFormError(
            "Informe a unidade escolar anterior para a qual o aluno deve o hist\u00f3rico."
        )

    return deve_historico_str == "sim", unidade_anterior


def build_dados_frequencia_form(form) -> dict:
    dados_frequencia = {"meses": []}
    algum_valido = False

    for mes_id, mes_nome in FREQUENCIA_FORM_MESES:
        dias_raw = (form.get(f"freq_{mes_id}_dias") or "").strip()
        faltas_raw = (form.get(f"freq_{mes_id}_faltas") or "").strip()

        if not dias_raw and not faltas_raw:
            dados_frequencia["meses"].append(
                {
                    "id": mes_id,
                    "nome": mes_nome,
                    "dias_letivos": None,
                    "faltas": None,
                    "frequencia": None,
                    "preenchido": False,
                }
            )
            continue

        try:
            dias = float(dias_raw.replace(",", ".")) if dias_raw else None
            faltas = float(faltas_raw.replace(",", ".")) if faltas_raw else None
        except ValueError:
            raise DeclaracaoFormError(
                "Verifique os valores de dias letivos e faltas informados na frequ\u00eancia."
            )

        if dias is None or faltas is None:
            raise DeclaracaoFormError(
                "Para cada m\u00eas de frequ\u00eancia preenchido, informe tanto os dias "
                "letivos quanto as faltas."
            )

        if dias <= 0 or faltas < 0 or faltas > dias:
            raise DeclaracaoFormError(
                "Os valores de dias letivos e faltas s\u00e3o inv\u00e1lidos em um ou mais meses. "
                "Verifique e tente novamente."
            )

        freq_percent = ((dias - faltas) / dias) * 100.0
        algum_valido = True

        dados_frequencia["meses"].append(
            {
                "id": mes_id,
                "nome": mes_nome,
                "dias_letivos": dias,
                "faltas": faltas,
                "frequencia": round(freq_percent, 1),
                "preenchido": True,
            }
        )

    if not algum_valido:
        raise DeclaracaoFormError(
            "Informe ao menos um m\u00eas de frequ\u00eancia com dias letivos e faltas v\u00e1lidos."
        )

    return dados_frequencia


def _format_frequency_number(value) -> str:
    try:
        number = float(value)
        if number.is_integer():
            return str(int(number))
        return f"{number:.1f}".replace(".", ",")
    except Exception:
        return str(value)


def _resolve_frequency_month(item, index) -> str:
    raw_month = item.get("nome_mes")
    if raw_month in (None, ""):
        raw_month = item.get("mes")
    if raw_month in (None, ""):
        raw_month = item.get("mes_nome")
    if raw_month in (None, ""):
        raw_month = item.get("descricao_mes")
    if raw_month in (None, ""):
        raw_month = item.get("descricao")

    month_name = ""

    if raw_month not in (None, ""):
        if isinstance(raw_month, (int, float)):
            month_index = int(raw_month)
            month_name = MESES_FREQUENCIA_BR.get(month_index, str(month_index))
        else:
            text = str(raw_month).strip()
            if text.isdigit():
                month_index = int(text)
                month_name = MESES_FREQUENCIA_BR.get(month_index, text)
            else:
                month_name = text

    if not month_name:
        month_name = MESES_FREQUENCIA_BR.get(index, f"M\u00eas {index}")

    return month_name


def _build_frequencia_tabela_html(meses) -> str:
    linhas_tabela = ""

    for index, item in enumerate(meses, start=1):
        nome_mes = _resolve_frequency_month(item, index)

        dias_val = item.get("dias_letivos_calculados")
        if dias_val is None:
            dias_val = item.get("dias_letivos")
        if dias_val is None:
            dias_val = item.get("dias")

        faltas_val = item.get("faltas_calculadas")
        if faltas_val is None:
            faltas_val = item.get("faltas")

        freq_val = item.get("frequencia")
        if freq_val is None:
            freq_val = item.get("freq")

        preenchido_raw = item.get("preenchido")
        if preenchido_raw is None:
            preenchido = any(
                value not in (None, "", 0, 0.0)
                for value in (dias_val, faltas_val, freq_val)
            )
        else:
            preenchido = bool(preenchido_raw)

        if preenchido:
            dias_txt = _format_frequency_number(dias_val) if dias_val not in (None, "") else "0"
            faltas_txt = _format_frequency_number(faltas_val) if faltas_val not in (None, "") else "0"
            if freq_val in (None, ""):
                freq_txt = "\u2014"
            else:
                try:
                    freq_txt = f"{float(freq_val):.1f}%".replace(".", ",")
                except Exception:
                    freq_txt = str(freq_val)
        else:
            dias_txt = "\u2014"
            faltas_txt = "\u2014"
            freq_txt = "\u2014"

        linhas_tabela += (
            "<tr>"
            f"<td style='border:1px solid #444;padding:4px 6px;text-align:center;'>{nome_mes}</td>"
            f"<td style='border:1px solid #444;padding:4px 6px;text-align:center;'>{dias_txt}</td>"
            f"<td style='border:1px solid #444;padding:4px 6px;text-align:center;'>{faltas_txt}</td>"
            f"<td style='border:1px solid #444;padding:4px 6px;text-align:center;'>{freq_txt}</td>"
            "</tr>"
        )

    return (
        "<br><br>"
        "<table style='width:75%;max-width:600px;margin:0 auto;"
        "border-collapse:collapse;font-size:12px;margin-top:4px;'>"
        "<thead>"
        "<tr>"
        "<th style='border:1px solid #444;padding:4px 6px;text-align:center;'>M\u00eas</th>"
        "<th style='border:1px solid #444;padding:4px 6px;text-align:center;'>Dias letivos</th>"
        "<th style='border:1px solid #444;padding:4px 6px;text-align:center;'>Faltas</th>"
        "<th style='border:1px solid #444;padding:4px 6px;text-align:center;'>Frequ\u00eancia</th>"
        "</tr>"
        "</thead>"
        "<tbody>"
        f"{linhas_tabela}"
        "</tbody>"
        "</table>"
        "<br>"
        "<span style='font-size:12px;color:#555;'>"
        "</span>"
    )


def _build_historico_unidade_html(unidade_anterior, escolas_df) -> str:
    if not unidade_anterior:
        return ""

    unidade_anterior = " ".join(str(unidade_anterior).strip().split())
    esc_df = None
    if escolas_df is not None:
        try:
            esc_df = escolas_df[
                escolas_df.iloc[:, 3].str.upper() == unidade_anterior.upper()
            ]
        except Exception:
            esc_df = None

    if esc_df is not None and not esc_df.empty:
        unidade_nome = esc_df.iloc[0, 3]
        inep = esc_df.iloc[0, 4]
        municipio = esc_df.iloc[0, 2]
        uf = esc_df.iloc[0, 1]
        return (
            "<div style='font-size:14px;'>"
            f"<strong>Unidade:</strong> {unidade_nome}<br>"
            f"<strong>INEP:</strong> {inep}<br>"
            f"<strong>Cidade:</strong> {municipio}<br>"
            f"<strong>Estado:</strong> {uf}<br><br>"
            "</div>"
        )

    return f"<strong>Unidade:</strong> {unidade_anterior}<br><br>"


def _build_historico_unidade_observacao(unidade_anterior, escolas_df) -> str:
    if not unidade_anterior:
        return ""

    unidade_anterior = " ".join(str(unidade_anterior).strip().split())
    esc_df = None
    if escolas_df is not None:
        try:
            esc_df = escolas_df[
                escolas_df.iloc[:, 3].str.upper() == unidade_anterior.upper()
            ]
        except Exception:
            esc_df = None

    if esc_df is not None and not esc_df.empty:
        unidade_nome = str(esc_df.iloc[0, 3]).strip()
        municipio = str(esc_df.iloc[0, 2]).strip()
        uf = str(esc_df.iloc[0, 1]).strip()
        local = f" - {municipio}/{uf}" if municipio or uf else ""
        return f"Unidade: {unidade_nome}{local}."

    return f"Unidade: {unidade_anterior}."


def _html(value) -> str:
    return escape(str(value or "").strip())


def _selected_key(tipo) -> str:
    if tipo == "Escolaridade":
        return "escolaridade"
    if tipo == "Transferencia":
        return "transferencia"
    if tipo == "Conclus\u00e3o":
        return "conclusao"
    return ""


def _ensino_label(segmento) -> str:
    if segmento == "EJA":
        return "segmento de Educa\u00e7\u00e3o de Jovens e Adultos (EJA)"
    return "Ensino Fundamental"


def _series_conclusao_text(serie, segmento, semestre_texto="") -> tuple[str, str]:
    if segmento == "EJA":
        mapping = {
            "1\u00aa S\u00c9RIE E.F": "2\u00aa S\u00c9RIE E.F",
            "2\u00aa S\u00c9RIE E.F": "3\u00aa S\u00c9RIE E.F",
            "3\u00aa S\u00c9RIE E.F": "4\u00aa S\u00c9RIE E.F",
            "4\u00aa S\u00c9RIE E.F": "5\u00aa S\u00c9RIE E.F",
            "5\u00aa S\u00c9RIE E.F": "6\u00aa S\u00c9RIE E.F",
            "6\u00aa S\u00c9RIE E.F": "7\u00aa S\u00c9RIE E.F",
            "7\u00aa S\u00c9RIE E.F": "8\u00aa S\u00c9RIE E.F",
            "8\u00aa S\u00c9RIE E.F": "1\u00aa S\u00c9RIE E.M",
            "1\u00aa S\u00c9RIE E.M": "2\u00aa S\u00c9RIE E.M",
            "2\u00aa S\u00c9RIE E.M": "3\u00aa S\u00c9RIE E.M",
            "3\u00aa S\u00c9RIE E.M": "ENSINO SUPERIOR",
        }
        serie_atual = str(serie)
        semestre = f" - {semestre_texto}" if semestre_texto else ""
        return f"{serie_atual}{semestre}", mapping.get(serie_atual.upper(), "a s\u00e9rie subsequente")

    match = re.search(r"(\d+)\u00ba", str(serie))
    next_year = int(match.group(1)) + 1 if match else None
    return str(serie), f"{next_year}\u00ba ano" if next_year else "a s\u00e9rie subsequente"


def _series_transferencia_text(serie) -> str:
    return re.sub(r"^(\d+\u00ba).*", r"\1 ano", str(serie))


def _build_siae_option(key, selected_key, text):
    selected = key == selected_key
    status = "Selecionada" if selected else "N\u00e3o preenchido"
    status_class = "selected" if selected else "locked"
    mark = "X" if selected else ""
    return (
        f'<div class="siae-option siae-option-{key} {status_class}" aria-label="{status}">'
        f'<span class="siae-check" aria-label="{status}">{mark}</span>'
        f'<div class="siae-option-text">{text}'
        "</div>"
        "</div>"
    )


def _build_siae_observations_html(observacoes):
    items = [str(obs or "").strip() for obs in observacoes if str(obs or "").strip()]
    if not items:
        return ""

    if len(items) == 1:
        return (
            '<section class="siae-observations bloco-observacoes">'
            f'<p><strong>Observa\u00e7\u00e3o:</strong> {_html(items[0])}</p>'
            "</section>"
        )

    paragraphs = "".join(
        f'<p class="siae-observation-paragraph paragrafo-observacao">{_html(item)}</p>'
        for item in items
    )
    return (
        '<section class="siae-observations bloco-observacoes">'
        '<p class="siae-observations-title"><strong>Observa\u00e7\u00f5es:</strong></p>'
        f'<div class="siae-observations-content conteudo-observacoes">{paragraphs}</div>'
        "</section>"
    )


def _build_modelo_siae_declaracao_html(
    *,
    tipo,
    segmento,
    nome,
    ra,
    ra_label,
    data_nasc,
    serie,
    horario="Desconhecido",
    semestre_texto="",
    ano=None,
    observacoes_adicionais=None,
):
    selected = _selected_key(tipo)
    if not selected:
        return ""

    ano = ano or datetime.now().year
    ensino = _ensino_label(segmento)
    serie_conclusao, serie_subsequente = _series_conclusao_text(serie, segmento, semestre_texto)
    serie_transferencia = _series_transferencia_text(serie)

    declaracao_intro = (
        '<div class="siae-student-lines">'
        '<div class="siae-student-line">'
        f'Declaro, para os devidos fins, que o(a) aluno(a) <span class="siae-fill long">{_html(nome)}</span>, '
        f'{_html(ra_label)} n\u00ba <span class="siae-fill medium">{_html(ra)}</span>, '
        f'nascido(a) em <span class="siae-fill short">{_html(data_nasc)}</span>, possui a seguinte situa\u00e7\u00e3o '
        'escolar nesta Unidade Escolar:'
        '</div>'
        '</div>'
    )

    conclusao = (
        f'Concluiu o(a) <span class="siae-fill small">{_html(serie_conclusao)}</span> do '
        f'<span class="siae-fill medium">{_html(ensino)}</span> nesta Unidade Escolar, no ano de '
        f'<span class="siae-fill tiny">{_html(ano)}</span>, estando apto(a) a ingressar no(a) '
        f'<span class="siae-fill small">{_html(serie_subsequente)}</span>.'
    )
    escolaridade = (
        f'\u00c9 aluno(a) regularmente matriculado(a) e frequente no(a) '
        f'<span class="siae-fill medium">{_html(serie)}</span> do '
        f'<span class="siae-fill medium">{_html(ensino)}</span> nesta Unidade Escolar.'
        f'<span class="siae-option-detail">Hor\u00e1rio de aula: {_html(horario)}.</span>'
    )
    transferencia = (
        'Pediu transfer\u00eancia nesta data e os documentos solicitados ser\u00e3o expedidos no prazo de '
        '<span class="siae-fill tiny">30 dias \u00fateis</span>. '
        'O(a) aluno(a) tem direito a matricular-se no(a) '
        f'<span class="siae-fill small">{_html(serie_transferencia)}</span> do '
        f'<span class="siae-fill medium">{_html(ensino)}</span>.'
    )

    observacoes = ["O Município adota o Ensino Fundamental de 09 anos."]
    observacoes.extend(observacoes_adicionais or [])
    observacoes_html = _build_siae_observations_html(observacoes)

    return (
        '<section class="siae-declaration" aria-label="Modelo municipal de declaração">'
        '<h2 class="siae-title">DECLARAÇÃO</h2>'
        f'{declaracao_intro}'
        '<div class="siae-options">'
        f'{_build_siae_option("conclusao", selected, conclusao)}'
        f'{_build_siae_option("escolaridade", selected, escolaridade)}'
        f'{_build_siae_option("transferencia", selected, transferencia)}'
        '</div>'
        f'{observacoes_html}'
        '<p class="siae-closing frase-final">Por ser expressão da verdade, firmamos a presente declaração.</p>'
        '</section>'
    )


def build_declaracao_escolar_context(
    *,
    tipo,
    segmento,
    nome,
    ra,
    ra_label,
    data_nasc,
    serie,
    horario="Desconhecido",
    semestre_texto="",
    row=None,
    notas_tabela_html="",
    deve_historico=False,
    unidade_anterior=None,
    escolas_df=None,
    dados_frequencia=None,
):
    """
    Monta titulo, texto e classes de impressao das declaracoes escolares.

    Retorna None para manter o contrato atual quando o tipo e desconhecido
    ou quando frequencia nao recebe meses.
    """
    is_eja = segmento == "EJA"
    declaracao_text = ""
    tem_observacoes = False
    observacoes_html = ""
    observacoes_siae = []

    if tipo == "Escolaridade":
        titulo = "Declara\u00e7\u00e3o de Escolaridade"
        if is_eja:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do {ra_label} "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, "
                f"encontra-se regularmente matriculado(a) no segmento de "
                f"<strong><u>Educa\u00e7\u00e3o de Jovens e Adultos (EJA)</u></strong> da "
                f"E.M Jos\u00e9 Padin Mouta, cursando atualmente o(a) "
                f"<strong><u>{serie}</u></strong>."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, "
                f"encontra-se regularmente matriculado(a) na "
                f"E.M Jos\u00e9 Padin Mouta, cursando atualmente o(a) "
                f"<strong><u>{serie}</u></strong> no hor\u00e1rio de aula: "
                f"<strong><u>{horario}</u></strong>."
            )

    elif tipo == "Transferencia":
        titulo = "Declara\u00e7\u00e3o de Transfer\u00eancia"
        if is_eja:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do {ra_label} "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, matriculado(a) no segmento de "
                f"<strong><u>Educa\u00e7\u00e3o de Jovens e Adultos (EJA)</u></strong> da "
                f"E.M Jos\u00e9 Padin Mouta, solicitou transfer\u00eancia desta unidade escolar "
                f"na data de hoje, estando apto(a) a cursar o(a) "
                f"<strong><u>{serie}</u></strong>."
            )

            if not deve_historico:
                declaracao_text += (
                    " Informamos, ainda, que o hist\u00f3rico escolar ser\u00e1 emitido no prazo de "
                    "at\u00e9 30 (trinta) dias."
                )
        else:
            serie_mod = re.sub(r"^(\d+\u00ba).*", r"\1 ano", str(serie))
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) respons\u00e1vel do(a) "
                f"aluno(a) <strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, compareceu a nossa "
                f"unidade escolar e solicitou transfer\u00eancia na data de hoje, "
                f"o aluno est\u00e1 apto(a) a cursar o(a) "
                f"<strong><u>{serie_mod}</u></strong>."
            )

            if not deve_historico:
                declaracao_text += (
                    " Informamos, ainda, que o hist\u00f3rico escolar ser\u00e1 emitido no prazo de "
                    "at\u00e9 30 (trinta) dias."
                )

            if notas_tabela_html:
                declaracao_text += notas_tabela_html

    elif tipo == "Conclus\u00e3o":
        titulo = "Declara\u00e7\u00e3o de Conclus\u00e3o"

        if is_eja:
            mapping = {
                "1\u00aa S\u00c9RIE E.F": "2\u00aa S\u00c9RIE E.F",
                "2\u00aa S\u00c9RIE E.F": "3\u00aa S\u00c9RIE E.F",
                "3\u00aa S\u00c9RIE E.F": "4\u00aa S\u00c9RIE E.F",
                "4\u00aa S\u00c9RIE E.F": "5\u00aa S\u00c9RIE E.F",
                "5\u00aa S\u00c9RIE E.F": "6\u00aa S\u00c9RIE E.F",
                "6\u00aa S\u00c9RIE E.F": "7\u00aa S\u00c9RIE E.F",
                "7\u00aa S\u00c9RIE E.F": "8\u00aa S\u00c9RIE E.F",
                "8\u00aa S\u00c9RIE E.F": "1\u00aa S\u00c9RIE E.M",
                "1\u00aa S\u00c9RIE E.M": "2\u00aa S\u00c9RIE E.M",
                "2\u00aa S\u00c9RIE E.M": "3\u00aa S\u00c9RIE E.M",
                "3\u00aa S\u00c9RIE E.M": "ENSINO SUPERIOR",
            }
            series_text = mapping.get(str(serie).upper(), "a s\u00e9rie subsequente")

            semestre_parte = (
                f", no <strong><u>{semestre_texto}</u></strong>"
                if semestre_texto
                else ""
            )

            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do {ra_label} "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, concluiu com \u00eaxito o(a) "
                f"<strong><u>{serie}</u></strong>{semestre_parte} no segmento de "
                f"<strong><u>Educa\u00e7\u00e3o de Jovens e Adultos (EJA)</u></strong> da "
                f"E.M Jos\u00e9 Padin Mouta, estando apto(a) a ingressar no(na) "
                f"<strong><u>{series_text}</u></strong>."
            )
        else:
            match = re.search(r"(\d+)\u00ba\s*ano", str(serie))
            next_year = int(match.group(1)) + 1 if match else None
            series_text = f"{next_year}\u00ba ano" if next_year else "a s\u00e9rie subsequente"

            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, concluiu com \u00eaxito o(a) "
                f"<strong><u>{serie}</u></strong>, estando apto(a) a ingressar no(na) "
                f"<strong><u>{series_text}</u></strong>."
            )

    elif tipo in ("Frequencia", "Frequ\u00eancia"):
        titulo = "Declara\u00e7\u00e3o de Frequ\u00eancia"

        if not dados_frequencia or not dados_frequencia.get("meses"):
            return None

        if is_eja:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do {ra_label} "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, regularmente matriculado(a) "
                f"no segmento de <strong><u>Educa\u00e7\u00e3o de Jovens e Adultos (EJA)</u></strong> "
                f"da E.M Jos\u00e9 Padin Mouta, teve sua frequ\u00eancia apurada nos meses abaixo "
                f"conforme quadro a seguir."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, regularmente matriculado(a) "
                f"no(a) <strong><u>{serie}</u></strong> da E.M Jos\u00e9 Padin Mouta, "
                f"teve sua frequ\u00eancia apurada nos meses abaixo conforme quadro a seguir."
            )

        declaracao_text += _build_frequencia_tabela_html(dados_frequencia.get("meses", []))

    else:
        return None

    row_data = row if row is not None else {}
    valor_bolsa = str(row_data.get("BOLSA FAMILIA", "")).strip().upper()

    if deve_historico or (valor_bolsa == "SIM" and tipo != "Escolaridade"):
        tem_observacoes = True
        observacoes_html += "<br><br><strong>Observa\u00e7\u00f5es:</strong><br>"
        observacoes_html += (
            '<label class="checkbox-label" '
            "style='display:block;text-align:justify;font-size:14px;'>"
        )

        if deve_historico:
            historico_observacao = "O aluno deve o histórico escolar da unidade anterior"
            unidade_observacao = _build_historico_unidade_observacao(
                unidade_anterior,
                escolas_df,
            )
            if unidade_observacao:
                historico_observacao += f", referente à {unidade_observacao.rstrip('.')}"
            historico_observacao += (
                ". Após sua entrega, o documento será confeccionado em até 30 dias úteis."
            )
            observacoes_siae.append(historico_observacao)

            observacoes_html += '<span class="warning-icon">&#9888;</span> '
            observacoes_html += (
                "O aluno deve o hist\u00f3rico escolar da unidade anterior:<br><br>"
            )
            observacoes_html += _build_historico_unidade_html(unidade_anterior, escolas_df)
            observacoes_html += (
                "Ap\u00f3s sua entrega, o documento ser\u00e1 confeccionado em "
                "at\u00e9 30 dias \u00fateis.<br><br>"
            )

        if valor_bolsa == "SIM" and tipo != "Escolaridade":
            observacoes_siae.append("O aluno é beneficiário do Programa Bolsa Família.")

            observacoes_html += (
                '<img src="/static/logos/bolsa_familia.jpg" '
                'alt="Bolsa Fam\u00edlia" '
                'style="width:28px;vertical-align:middle;margin-right:5px;">'
                "O aluno \u00e9 benefici\u00e1rio do Programa Bolsa Fam\u00edlia."
            )

        observacoes_html += "</label>"
        declaracao_text += observacoes_html

    body_classes = []
    if tipo in ("Frequencia", "Frequ\u00eancia"):
        body_classes.append("tipo-frequencia")
    if tipo == "Transferencia" and tem_observacoes:
        body_classes.append("transferencia-com-observacoes")

    if tipo in ("Escolaridade", "Transferencia", "Conclus\u00e3o"):
        declaracao_text = _build_modelo_siae_declaracao_html(
            tipo=tipo,
            segmento=segmento,
            nome=nome,
            ra=ra,
            ra_label=ra_label,
            data_nasc=data_nasc,
            serie=serie,
            horario=horario,
            semestre_texto=semestre_texto,
            observacoes_adicionais=observacoes_siae,
        )
        if tipo == "Transferencia" and notas_tabela_html:
            declaracao_text += '<div class="declaracao-extra declaracao-notas">'
            declaracao_text += notas_tabela_html
            declaracao_text += "</div>"
        body_classes.insert(0, "modelo-siae")

    return {
        "titulo": titulo,
        "declaracao_text": declaracao_text,
        "body_classes": body_classes,
    }


def build_declaracao_personalizada_context(dados):
    """
    Monta titulo e texto HTML das declaracoes personalizadas.

    Retorna None quando o tipo de declaracao nao e reconhecido.
    """
    nome = get_str(dados, "nome_aluno")
    ra = get_str(dados, "ra")
    data_nasc = parse_data_nascimento_personalizada(get_str(dados, "data_nascimento"))
    segmento = normalizar_segmento_personalizado(dados)
    segmento_label, prep_segmento = contexto_segmento(segmento)
    tipo_decl = normalizar_tipo_declaracao(dados)

    declaracao_text = ""
    titulo = ""

    if tipo_decl in ("conclusao", "conclus\u00e3o"):
        titulo = "Declara\u00e7\u00e3o de Conclus\u00e3o"
        ano_serie = get_str(dados, "ano_serie_concluida")
        ano_conclusao = get_str(dados, "ano_conclusao")

        deve_hist_value = dados.get("deve_historico_unidade")
        deve_hist_text = str(deve_hist_value or "").strip().lower()
        deve_hist_unidade = deve_hist_text in ("sim", "1", "true", "on")

        if segmento == "Fundamental":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, concluiu o(a) "
                f"<strong><u>{ano_serie}</u></strong> {prep_segmento} "
                f"<strong><u>{segmento_label}</u></strong>, no ano letivo de "
                f"<strong><u>{ano_conclusao}</u></strong>, nesta unidade escolar."
            )
        else:
            semestre_conclusao = normalizar_semestre(
                dados,
                "semestre_conclusao",
                "semestre_conclusao_opcao",
                "semestre_matricula",
                "semestre_matricula_opcao",
            )

            if semestre_conclusao:
                periodo_conclusao = (
                    f"no <strong><u>{semestre_conclusao}</u></strong> do ano de "
                    f"<strong><u>{ano_conclusao}</u></strong>"
                )
            else:
                periodo_conclusao = f"no ano letivo de <strong><u>{ano_conclusao}</u></strong>"

            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, concluiu o(a) "
                f"<strong><u>{ano_serie}</u></strong> no segmento de "
                f"<strong><u>Educa\u00e7\u00e3o de Jovens e Adultos (EJA)</u></strong>, "
                f"{periodo_conclusao}, nesta unidade escolar."
            )

        if deve_hist_unidade:
            declaracao_text += (
                " Ressalta-se que consta, junto a esta unidade escolar, "
                "pend\u00eancia de hist\u00f3rico escolar referente ao(\u00e0) aluno(a) citado(a)."
            )

    elif tipo_decl in ("matriculacancelada", "matricula cancelada", "matricula_cancelada"):
        titulo = "Declara\u00e7\u00e3o de Matr\u00edcula Cancelada"
        ano_serie = get_str(dados, "ano_serie_matricula")
        ano_matricula = get_str(dados, "ano_matricula")
        semestre_matricula = normalizar_semestre(
            dados,
            "semestre_matricula",
            "semestre_matricula_opcao",
        )

        if segmento == "EJA" and semestre_matricula:
            periodo_matricula = (
                f"no <strong><u>{semestre_matricula}</u></strong> do ano de "
                f"<strong><u>{ano_matricula}</u></strong>"
            )
        else:
            periodo_matricula = f"no ano de <strong><u>{ano_matricula}</u></strong>"

        if segmento == "EJA":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, esteve matriculado(a) no(a) "
                f"<strong><u>{ano_serie}</u></strong> no segmento de "
                f"<strong><u>Educa\u00e7\u00e3o de Jovens e Adultos (EJA)</u></strong>, "
                f"{periodo_matricula}, nesta unidade escolar, tendo sua matr\u00edcula cancelada."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, esteve matriculado(a) no(a) "
                f"<strong><u>{ano_serie}</u></strong> {prep_segmento} "
                f"<strong><u>{segmento_label}</u></strong>, {periodo_matricula}, "
                "nesta unidade escolar, tendo sua matr\u00edcula cancelada."
            )

    elif tipo_decl == "ncom":
        titulo = "Declara\u00e7\u00e3o de N\u00e3o Comparecimento (NCOM)"
        ano_serie = get_str(dados, "ano_serie_vaga")
        ano_ref = get_str(dados, "ano_referencia_ncom")
        semestre_ref = normalizar_semestre(
            dados,
            "semestre_referencia_ncom",
            "semestre_referencia",
        )

        if segmento == "EJA" and semestre_ref:
            periodo_ref = (
                f"para o <strong><u>{semestre_ref}</u></strong> do ano de "
                f"<strong><u>{ano_ref}</u></strong>"
            )
        else:
            periodo_ref = f"para o ano de <strong><u>{ano_ref}</u></strong>"

        if segmento == "EJA":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, teve vaga destinada ao(\u00e0) "
                f"<strong><u>{ano_serie}</u></strong> no segmento de "
                f"<strong><u>Educa\u00e7\u00e3o de Jovens e Adultos (EJA)</u></strong>, "
                f"{periodo_ref} nesta unidade escolar. Todavia, o(a) aluno(a) "
                "n\u00e3o compareceu \u00e0 unidade escolar, sendo considerado(a) NCOM \u2013 "
                "N\u00e3o Comparecimento, motivo pelo qual a vaga foi cancelada nesta "
                "unidade escolar."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, teve vaga destinada ao(\u00e0) "
                f"<strong><u>{ano_serie}</u></strong> {prep_segmento} "
                f"<strong><u>{segmento_label}</u></strong>, {periodo_ref} "
                "nesta unidade escolar. Todavia, o(a) aluno(a) n\u00e3o compareceu \u00e0 unidade "
                "escolar, sendo considerado(a) NCOM \u2013 N\u00e3o Comparecimento, motivo pelo qual "
                "a vaga foi cancelada nesta unidade escolar."
            )
    else:
        return None

    return {
        "titulo": titulo,
        "declaracao_text": declaracao_text,
    }


def _format_nota(value):
    if pd.isna(value):
        return "\u2014", None

    text = str(value).strip()
    if text == "":
        return "\u2014", None

    try:
        number = float(text.replace(",", "."))
    except Exception:
        return text, None

    color = "red" if number < 5 else "blue"
    return f"{number:.2f}".replace(".", ","), color


def build_notas_tabela_html(file_path, rm_num) -> str:
    try:
        notas_df = pd.read_excel(file_path, sheet_name="NOTAS")
        notas_df.columns = [str(col).strip().upper() for col in notas_df.columns]

        rm_col = None
        for col in notas_df.columns:
            if str(col).strip().upper() == "RM":
                rm_col = col
                break
        if rm_col is None and len(notas_df.columns) >= 3:
            rm_col = notas_df.columns[2]

        if rm_col is None:
            return ""

        notas_df = notas_df.copy()
        notas_df.loc[:, "RM_str"] = notas_df[rm_col].apply(format_rm)
        notas_aluno = notas_df[notas_df["RM_str"] == rm_num]
        if notas_aluno.empty:
            return ""

        notas_row = notas_aluno.iloc[0]
        linhas_notas = ""

        for nome_disc, col_1, col_2, col_3 in NOTAS_MATERIAS:
            n1_txt, n1_cor = _format_nota(notas_row.get(col_1))
            n2_txt, n2_cor = _format_nota(notas_row.get(col_2))
            n3_txt, n3_cor = _format_nota(notas_row.get(col_3))

            style_n1 = "border:1px solid #444;padding:3px 6px;text-align:center;"
            style_n2 = "border:1px solid #444;padding:3px 6px;text-align:center;"
            style_n3 = "border:1px solid #444;padding:3px 6px;text-align:center;"

            if n1_cor:
                style_n1 += f"color:{n1_cor};"
            if n2_cor:
                style_n2 += f"color:{n2_cor};"
            if n3_cor:
                style_n3 += f"color:{n3_cor};"

            linhas_notas += (
                "<tr>"
                f"<td style='border:1px solid #444;padding:3px 6px;text-align:left;'>{nome_disc}</td>"
                f"<td style='{style_n1}'>{n1_txt}</td>"
                f"<td style='{style_n2}'>{n2_txt}</td>"
                f"<td style='{style_n3}'>{n3_txt}</td>"
                "</tr>"
            )

        if not linhas_notas:
            return ""

        return (
            "<br>"
            "<span style='font-size:12px;'>"
            "<strong>As notas do aluno, por componente curricular, s\u00e3o:</strong>"
            "</span>"
            "<br>"
            "<table style='width:85%;max-width:700px;margin:4px auto 0 auto;"
            "border-collapse:collapse;font-size:11px;'>"
            "<thead>"
            "<tr>"
            "<th style='border:1px solid #444;padding:3px 6px;text-align:left;'>Componente curricular</th>"
            "<th style='border:1px solid #444;padding:3px 6px;text-align:center;'>1\u00ba trim.</th>"
            "<th style='border:1px solid #444;padding:3px 6px;text-align:center;'>2\u00ba trim.</th>"
            "<th style='border:1px solid #444;padding:3px 6px;text-align:center;'>3\u00ba trim.</th>"
            "</tr>"
            "</thead>"
            "<tbody>"
            f"{linhas_notas}"
            "</tbody>"
            "</table>"
        )
    except Exception as exc:
        print(f"[ERRO] Falha ao carregar notas do aluno (Transfer\u00eancia Fundamental): {exc}")
        return ""
