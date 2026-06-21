import tempfile
import unittest
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd

from services.declaracoes import (
    DeclaracaoFormError,
    build_declaracao_personalizada_payload,
    build_declaracao_escolar_context,
    build_lote_conclusao_5ano_context,
    build_lote_escolaridade_5ano_context,
    build_declaracao_personalizada_context,
    build_dados_frequencia_form,
    build_notas_tabela_html,
    contexto_segmento,
    data_extenso_praia_grande,
    format_date_br,
    format_eja_rm,
    format_rm,
    format_serie_ano,
    list_declaracao_alunos,
    load_declaracao_aluno_context,
    normalizar_segmento_personalizado,
    normalizar_semestre,
    normalizar_tipo_declaracao,
    normalizar_tipo_escolar_form,
    parse_data_nascimento_personalizada,
    read_lista_corrida_fundamental,
    read_lista_eja_declaracoes,
    resolve_historico_fields,
)


class DeclaracoesServiceTests(unittest.TestCase):
    def test_format_helpers_preserve_existing_rules(self):
        self.assertEqual(format_rm("00123.0"), "123")
        self.assertEqual(format_rm("ABC"), "ABC")
        self.assertEqual(format_eja_rm(123.0), "123")
        self.assertEqual(format_eja_rm(0), "")

        self.assertEqual(format_date_br("2026-05-25"), "25/05/2026")
        self.assertEqual(format_date_br("invalida"), "Desconhecida")
        self.assertEqual(format_serie_ano("5\u00baA"), "5\u00ba ano A")
        self.assertEqual(data_extenso_praia_grande(datetime(2026, 5, 25)), "Praia Grande, 25 de Maio de 2026")

    def test_personalizada_normalizers(self):
        dados = {
            "segmento_personalizado": "ef",
            "tipo_declaracao_personalizada": " Conclus\u00e3o ",
            "semestre_conclusao": "",
            "semestre_matricula": "1\u00ba semestre",
        }

        self.assertEqual(normalizar_segmento_personalizado(dados), "Fundamental")
        self.assertEqual(contexto_segmento("Fundamental"), ("Ensino Fundamental", "do"))
        self.assertEqual(normalizar_tipo_declaracao(dados), "conclus\u00e3o")
        self.assertEqual(normalizar_semestre(dados, "semestre_conclusao", "semestre_matricula"), "1\u00ba semestre")
        self.assertEqual(parse_data_nascimento_personalizada("2026-05-25"), "25/05/2026")
        self.assertEqual(parse_data_nascimento_personalizada("25/05/2026"), "25/05/2026")
        self.assertEqual(parse_data_nascimento_personalizada(""), "Desconhecida")

    def test_declaracao_form_helpers_personalizada_payload(self):
        payload = build_declaracao_personalizada_payload(
            {
                "segmento_personalizado": "EJA",
                "nome_aluno": "Aluno EJA",
                "data_nascimento": "2026-05-25",
                "ra": "123",
                "tipo_declaracao_personalizada": "Conclusao",
                "ano_serie_concluida": "8\u00aa S\u00c9RIE E.F",
                "ano_conclusao": "2026",
                "deve_historico_unidade": "N\u00e3o",
                "semestre_conclusao": "2\u00ba semestre",
            }
        )

        self.assertEqual(payload["segmento"], "EJA")
        self.assertFalse(payload["deve_historico_unidade"])
        self.assertEqual(payload["semestre_conclusao"], "2\u00ba semestre")

        with self.assertRaises(DeclaracaoFormError):
            build_declaracao_personalizada_payload(
                {
                    "segmento_personalizado": "EJA",
                    "nome_aluno": "Aluno EJA",
                    "data_nascimento": "2026-05-25",
                    "ra": "123",
                    "tipo_declaracao_personalizada": "Conclusao",
                    "ano_serie_concluida": "8\u00aa S\u00c9RIE E.F",
                    "ano_conclusao": "2026",
                    "deve_historico_unidade": "N\u00e3o",
                }
            )

    def test_declaracao_form_helpers_tipo_historico_e_frequencia(self):
        self.assertEqual(normalizar_tipo_escolar_form("transfer\u00eancia"), "Transferencia")
        self.assertEqual(normalizar_tipo_escolar_form("Conclusao"), "Conclus\u00e3o")
        self.assertEqual(normalizar_tipo_escolar_form("frequ\u00eancia"), "Frequencia")

        deve, unidade = resolve_historico_fields(
            "Transferencia",
            "sim",
            unidade_select=" Escola A ",
            unidade_manual="",
        )
        self.assertTrue(deve)
        self.assertEqual(unidade, "Escola A")

        deve, unidade = resolve_historico_fields("Escolaridade", None, "", "Escola B")
        self.assertFalse(deve)
        self.assertEqual(unidade, "")

        frequencia = build_dados_frequencia_form(
            {
                "freq_jan_dias": "20",
                "freq_jan_faltas": "2",
            }
        )
        self.assertEqual(frequencia["meses"][0]["frequencia"], 90.0)
        self.assertTrue(frequencia["meses"][0]["preenchido"])
        self.assertFalse(frequencia["meses"][1]["preenchido"])

        with self.assertRaises(DeclaracaoFormError):
            build_dados_frequencia_form({"freq_jan_dias": "20"})

    def test_build_notas_tabela_html_formats_values_and_colors(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir, "notas.xlsx")
            notas = pd.DataFrame(
                [
                    {
                        "RM": 123,
                        "LP_1T": 4.5,
                        "LP_2T": 8,
                        "LP_3T": "",
                    }
                ]
            )
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                notas.to_excel(writer, sheet_name="NOTAS", index=False)

            html = build_notas_tabela_html(path, "123")

        self.assertIn("L\u00edngua Portuguesa", html)
        self.assertIn("4,50", html)
        self.assertIn("color:red", html)
        self.assertIn("8,00", html)
        self.assertIn("color:blue", html)

    def test_load_declaracao_aluno_context_fundamental_from_lista_corrida(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir, "lista_fundamental.xlsx")
            lista = pd.DataFrame(
                [
                    {
                        "RM": 123,
                        "NOME": "Aluno Fundamental",
                        "S\u00c9RIE": "5\u00baA",
                        "DATA NASC.": "2020-01-02",
                        "RA": "RA123",
                        "HOR\u00c1RIO": "13h00",
                        "BOLSA FAMILIA": "NAO",
                    }
                ]
            )
            notas = pd.DataFrame([{"RM": 123, "LP_1T": 8, "LP_2T": 9, "LP_3T": 10}])
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                lista.to_excel(writer, sheet_name="LISTA CORRIDA", index=False)
                notas.to_excel(writer, sheet_name="NOTAS", index=False)

            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter("always", FutureWarning)
                planilha = read_lista_corrida_fundamental(path)
                context = load_declaracao_aluno_context(
                    path,
                    "123",
                    "Fundamental",
                    "Transferencia",
                )
                alunos = list_declaracao_alunos(path, "Fundamental")

        future_warnings = [
            warning for warning in caught if issubclass(warning.category, FutureWarning)
        ]
        self.assertEqual(future_warnings, [])
        self.assertEqual(planilha.iloc[0]["RM_str"], "123")
        self.assertEqual(context["nome"], "Aluno Fundamental")
        self.assertEqual(context["serie"], "5\u00ba ano A")
        self.assertEqual(context["data_nasc"], "02/01/2020")
        self.assertEqual(context["notas_tabela_html"], "")
        self.assertEqual(alunos, [{"rm": "123", "nome": "Aluno Fundamental"}])

    def test_load_declaracao_aluno_context_eja_from_lista(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir, "lista_eja.xlsx")
            ignored_header = [None] * 30
            aluno = [""] * 30
            aluno[0] = "8\u00aa S\u00c9RIE E.F"
            aluno[2] = 456
            aluno[3] = "Aluno EJA"
            aluno[6] = "2020-02-03"
            aluno[7] = 0
            aluno[8] = "RG456"
            aluno[29] = "2\u00ba semestre"
            pd.DataFrame([ignored_header, aluno]).to_excel(
                path,
                index=False,
                header=False,
            )

            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter("always", FutureWarning)
                df = read_lista_eja_declaracoes(path)
                context = load_declaracao_aluno_context(path, "456", "EJA", "Conclus\u00e3o")
                alunos = list_declaracao_alunos(path, "EJA")

        future_warnings = [
            warning for warning in caught if issubclass(warning.category, FutureWarning)
        ]
        self.assertEqual(future_warnings, [])
        self.assertEqual(df.iloc[0]["RM_str"], "456")
        self.assertEqual(context["nome"], "Aluno EJA")
        self.assertEqual(context["ra"], "RG456")
        self.assertEqual(context["ra_label"], "RG")
        self.assertEqual(context["data_nasc"], "03/02/2020")
        self.assertEqual(context["semestre_texto"], "2\u00ba semestre")
        self.assertEqual(alunos, [{"rm": "456", "nome": "Aluno EJA"}])

    def test_build_lote_escolaridade_5ano_context_filters_and_formats(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir, "lista_fundamental.xlsx")
            lista = pd.DataFrame(
                [
                    {
                        "RM": 123,
                        "NOME": "Aluno Quinto",
                        "S\u00c9RIE": "5\u00baA",
                        "DATA NASC.": "2020-01-02",
                        "RA": "RA123",
                        "HOR\u00c1RIO": "13h00",
                        "BOLSA FAMILIA": "NAO",
                    },
                    {
                        "RM": 456,
                        "NOME": "Aluno Quarto",
                        "S\u00c9RIE": "4\u00baA",
                        "DATA NASC.": "2020-01-02",
                        "RA": "RA456",
                        "HOR\u00c1RIO": "",
                        "BOLSA FAMILIA": "NAO",
                    },
                ]
            )
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                lista.to_excel(writer, sheet_name="LISTA CORRIDA", index=False)

            context = build_lote_escolaridade_5ano_context(
                path,
                now=datetime(2026, 5, 25),
            )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de Escolaridade")
        self.assertEqual(context["data_extenso"], "Praia Grande, 25 de Maio de 2026")
        self.assertEqual(len(context["registros"]), 1)
        registro = context["registros"][0]
        self.assertEqual(registro["nome"], "Aluno Quinto")
        self.assertEqual(registro["serie_fmt"], "5\u00ba ano A")
        self.assertIn("hor\u00e1rio de aula", registro["texto"])
        self.assertIn("13h00", registro["texto"])

    def test_build_lote_conclusao_5ano_context_duas_vias_and_bolsa_familia(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir, "lista_fundamental.xlsx")
            lista = pd.DataFrame(
                [
                    {
                        "RM": 123,
                        "NOME": "Aluno Conclusao",
                        "S\u00c9RIE": "5\u00baA",
                        "DATA NASC.": "2020-01-02",
                        "RA": "RA123",
                        "HOR\u00c1RIO": "13h00",
                        "BOLSA FAMILIA": "SIM",
                    }
                ]
            )
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                lista.to_excel(writer, sheet_name="LISTA CORRIDA", index=False)

            context = build_lote_conclusao_5ano_context(
                path,
                "Praia Grande, 22 de dezembro de 2026",
                bolsa_familia_src="/static/teste.jpg",
            )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de Conclus\u00e3o")
        self.assertEqual(context["total"], 2)
        self.assertEqual([registro["via"] for registro in context["registros"]], [1, 2])
        self.assertEqual(context["registros"][0]["series_text"], "6\u00ba ano")
        self.assertIn("Bolsa Fam\u00edlia", context["registros"][0]["texto"])
        self.assertIn('/static/teste.jpg', context["registros"][0]["texto"])

    def test_build_declaracao_escolar_context_fundamental_escolaridade(self):
        context = build_declaracao_escolar_context(
            tipo="Escolaridade",
            segmento="Fundamental",
            nome="Aluno Fundamental",
            ra="123",
            ra_label="RA",
            data_nasc="25/05/2026",
            serie="5\u00ba ano A",
            horario="13h00 as 17h00",
            row={},
        )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de Escolaridade")
        self.assertIn("Aluno Fundamental", context["declaracao_text"])
        self.assertIn("possui a seguinte situa\u00e7\u00e3o escolar", context["declaracao_text"])
        self.assertIn("Hor\u00e1rio de aula", context["declaracao_text"])
        self.assertIn("regularmente matriculado(a) e frequente", context["declaracao_text"])
        self.assertIn("Observa\u00e7\u00e3o:", context["declaracao_text"])
        self.assertIn("siae-observations", context["declaracao_text"])
        self.assertIn("Por ser express\u00e3o da verdade", context["declaracao_text"])
        self.assertIn("siae-option-escolaridade selected", context["declaracao_text"])
        self.assertIn("siae-student-line", context["declaracao_text"])
        self.assertIn("N\u00e3o preenchido", context["declaracao_text"])
        self.assertNotIn("na cidade de", context["declaracao_text"])
        self.assertNotIn("siae-municipal-heading", context["declaracao_text"])
        self.assertEqual(context["body_classes"], ["modelo-siae"])

    def test_build_declaracao_escolar_context_transferencia_observacoes(self):
        escolas_df = pd.DataFrame(
            [
                ["", "SP", "Praia Grande", "ESCOLA ORIGEM", "123456"],
            ]
        )

        context = build_declaracao_escolar_context(
            tipo="Transferencia",
            segmento="Fundamental",
            nome="Aluno Transferencia",
            ra="456",
            ra_label="RA",
            data_nasc="25/05/2026",
            serie="4\u00ba ano B",
            row={"BOLSA FAMILIA": "SIM"},
            notas_tabela_html="<table><tr><td>Notas</td></tr></table>",
            deve_historico=True,
            unidade_anterior="  ESCOLA   ORIGEM ",
            escolas_df=escolas_df,
        )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de Transfer\u00eancia")
        self.assertIn("4\u00ba ano", context["declaracao_text"])
        self.assertNotIn("Notas", context["declaracao_text"])
        self.assertIn("ESCOLA ORIGEM", context["declaracao_text"])
        self.assertIn("Bolsa Fam\u00edlia", context["declaracao_text"])
        self.assertIn("Observa\u00e7\u00f5es:", context["declaracao_text"])
        self.assertIn("O Munic\u00edpio adota o Ensino Fundamental de 09 anos.", context["declaracao_text"])
        self.assertIn(
            "O aluno deve o hist\u00f3rico escolar da unidade anterior, referente \u00e0 Unidade: "
            "ESCOLA ORIGEM - Praia Grande/SP. Ap\u00f3s sua entrega, o documento ser\u00e1 confeccionado "
            "em at\u00e9 30 dias \u00fateis.",
            context["declaracao_text"],
        )
        self.assertIn("paragrafo-observacao", context["declaracao_text"])
        self.assertNotIn("<ul>", context["declaracao_text"])
        self.assertNotIn("<li>", context["declaracao_text"])
        self.assertEqual(context["declaracao_text"].count("Observa\u00e7\u00e3o:"), 0)
        self.assertEqual(context["declaracao_text"].count("Observa\u00e7\u00f5es:"), 1)
        self.assertIn("siae-option-transferencia selected", context["declaracao_text"])
        self.assertEqual(context["body_classes"], ["modelo-siae", "transferencia-com-observacoes"])

    def test_build_declaracao_escolar_context_eja_conclusao(self):
        context = build_declaracao_escolar_context(
            tipo="Conclus\u00e3o",
            segmento="EJA",
            nome="Aluno EJA",
            ra="RG-1",
            ra_label="RG",
            data_nasc="25/05/2026",
            serie="8\u00aa S\u00c9RIE E.F",
            semestre_texto="2\u00ba semestre",
            row={},
        )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de Conclus\u00e3o")
        self.assertIn("2\u00ba semestre", context["declaracao_text"])
        self.assertIn("1\u00aa S\u00c9RIE E.M", context["declaracao_text"])
        self.assertIn("siae-option-conclusao selected", context["declaracao_text"])
        self.assertEqual(context["body_classes"], ["modelo-siae"])

    def test_build_declaracao_escolar_context_frequencia_and_invalid(self):
        context = build_declaracao_escolar_context(
            tipo="Frequ\u00eancia",
            segmento="Fundamental",
            nome="Aluno Frequencia",
            ra="789",
            ra_label="RA",
            data_nasc="25/05/2026",
            serie="3\u00ba ano C",
            row={},
            dados_frequencia={
                "meses": [
                    {
                        "mes": 2,
                        "dias_letivos": 20,
                        "faltas": 1,
                        "frequencia": 95,
                        "preenchido": True,
                    },
                    {"descricao": "Mar\u00e7o", "preenchido": False},
                ]
            },
        )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de Frequ\u00eancia")
        self.assertIn("Fevereiro", context["declaracao_text"])
        self.assertIn("95,0%", context["declaracao_text"])
        self.assertIn("Mar\u00e7o", context["declaracao_text"])
        self.assertEqual(context["body_classes"], ["tipo-frequencia"])
        self.assertIsNone(
            build_declaracao_escolar_context(
                tipo="Frequencia",
                segmento="Fundamental",
                nome="Aluno",
                ra="1",
                ra_label="RA",
                data_nasc="25/05/2026",
                serie="1\u00ba ano",
                row={},
                dados_frequencia={},
            )
        )
        self.assertIsNone(
            build_declaracao_escolar_context(
                tipo="Desconhecida",
                segmento="Fundamental",
                nome="Aluno",
                ra="1",
                ra_label="RA",
                data_nasc="25/05/2026",
                serie="1\u00ba ano",
                row={},
            )
        )

    def test_build_declaracao_personalizada_context_conclusao(self):
        context = build_declaracao_personalizada_context(
            {
                "segmento": "Fundamental",
                "tipo_declaracao": "Conclus\u00e3o",
                "nome_aluno": "Aluno Ficticio",
                "ra": "123",
                "data_nascimento": "2026-05-25",
                "ano_serie_concluida": "5\u00ba ano",
                "ano_conclusao": "2026",
                "deve_historico_unidade": "sim",
            }
        )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de Conclus\u00e3o")
        self.assertIn("Aluno Ficticio", context["declaracao_text"])
        self.assertIn("Ensino Fundamental", context["declaracao_text"])
        self.assertIn("pend\u00eancia de hist\u00f3rico", context["declaracao_text"])

    def test_build_declaracao_personalizada_context_matricula_cancelada_eja(self):
        context = build_declaracao_personalizada_context(
            {
                "segmento": "EJA",
                "tipo_declaracao": "matricula_cancelada",
                "nome_aluno": "Aluno EJA",
                "ra": "456",
                "data_nascimento": "25/05/2026",
                "ano_serie_matricula": "2\u00aa S\u00c9RIE E.F",
                "ano_matricula": "2026",
                "semestre_matricula": "1\u00ba semestre",
            }
        )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de Matr\u00edcula Cancelada")
        self.assertIn("Educa\u00e7\u00e3o de Jovens e Adultos", context["declaracao_text"])
        self.assertIn("1\u00ba semestre", context["declaracao_text"])

    def test_build_declaracao_personalizada_context_ncom_and_invalid(self):
        context = build_declaracao_personalizada_context(
            {
                "segmento": "Fundamental",
                "tipo_declaracao": "ncom",
                "nome_aluno": "Aluno NCOM",
                "ra": "789",
                "data_nascimento": "2026-05-25",
                "ano_serie_vaga": "4\u00ba ano",
                "ano_referencia_ncom": "2026",
            }
        )

        self.assertEqual(context["titulo"], "Declara\u00e7\u00e3o de N\u00e3o Comparecimento (NCOM)")
        self.assertIn("N\u00e3o Comparecimento", context["declaracao_text"])
        self.assertIsNone(build_declaracao_personalizada_context({"tipo_declaracao": "desconhecida"}))


if __name__ == "__main__":
    unittest.main()
