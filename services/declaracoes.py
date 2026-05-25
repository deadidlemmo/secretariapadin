import re
from datetime import datetime

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
.print-button {
  background-color: #283E51;
  color: #fff;
  border: none;
  padding: 10px 20px;
  border-radius: 5px;
  cursor: pointer;
  margin-top: 20px;
}
.print-button:hover {
  background-color: #1d2d3a;
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
