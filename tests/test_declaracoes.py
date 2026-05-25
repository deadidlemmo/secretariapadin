import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import pandas as pd

from services.declaracoes import (
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


if __name__ == "__main__":
    unittest.main()
