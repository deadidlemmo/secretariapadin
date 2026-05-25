import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import pandas as pd

from services.declaracoes import (
    build_declaracao_escolar_context,
    build_declaracao_personalizada_context,
    build_notas_tabela_html,
    contexto_segmento,
    data_extenso_praia_grande,
    format_date_br,
    format_eja_rm,
    format_rm,
    format_serie_ano,
    normalizar_segmento_personalizado,
    normalizar_semestre,
    normalizar_tipo_declaracao,
    parse_data_nascimento_personalizada,
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
        self.assertIn("hor\u00e1rio de aula", context["declaracao_text"])
        self.assertEqual(context["body_classes"], [])

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
        self.assertIn("Notas", context["declaracao_text"])
        self.assertIn("ESCOLA ORIGEM", context["declaracao_text"])
        self.assertIn("Bolsa Fam\u00edlia", context["declaracao_text"])
        self.assertEqual(context["body_classes"], ["transferencia-com-observacoes"])

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
