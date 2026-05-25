import unittest

from openpyxl import Workbook

from services.quadros_transferencias import (
    add_transfer_alerts_sheet,
    label_set,
    normalize_tipo_te,
    push_missing_info_alert,
    serie_key_from_value,
)


class QuadrosTransferenciasServiceTests(unittest.TestCase):
    def test_push_missing_info_alert_deduplicates_and_fills_blanks(self):
        alerts = []
        seen = set()

        payload = {
            "turma": "5A",
            "nome": "Aluno Ficticio",
            "ra": "",
            "tipo": "Sem Informacao",
            "data_str": "",
            "campo": "Tipo",
            "detalhe": "Campo vazio",
        }
        push_missing_info_alert(alerts, seen, **payload)
        push_missing_info_alert(alerts, seen, **payload)

        self.assertEqual(len(alerts), 1)
        self.assertEqual(alerts[0]["turma"], "5A")
        self.assertEqual(alerts[0]["ra"], "-")
        self.assertEqual(alerts[0]["data"], "-")

    def test_add_transfer_alerts_sheet_creates_alert_tab(self):
        wb = Workbook()

        add_transfer_alerts_sheet(
            wb,
            [
                {
                    "turma": "4A",
                    "nome": "Aluno Ficticio",
                    "ra": "123",
                    "tipo": "Rede Estadual",
                    "data": "01/02/2026",
                    "campo": "TIPO TE",
                    "detalhe": "Validado",
                }
            ],
        )

        self.assertIn("ALERTAS", wb.sheetnames)
        ws = wb["ALERTAS"]
        self.assertEqual(ws["A1"].value, "Turma")
        self.assertEqual(ws["B2"].value, "Aluno Ficticio")

    def test_label_set_preserves_template_prefix(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "Responsavel: antigo"

        label_set(ws, "A1", "Responsavel", "Novo")

        self.assertEqual(ws["A1"].value, "Responsavel: Novo")

    def test_serie_key_from_value_accepts_expected_series(self):
        self.assertEqual(serie_key_from_value("4\u00baF"), "4\u00ba")
        self.assertEqual(serie_key_from_value("5o D"), "5\u00ba")
        self.assertIsNone(serie_key_from_value("1\u00baA"))
        self.assertIsNone(serie_key_from_value("2019"))
        self.assertIsNone(serie_key_from_value(""))

    def test_normalize_tipo_te_keeps_current_labels(self):
        self.assertEqual(normalize_tipo_te("rede estadual"), "Rede Estadual")
        self.assertEqual(normalize_tipo_te("Sao Paulo"), "S\u00e3o Paulo")
        self.assertEqual(normalize_tipo_te("pais"), "Pa\u00eds")
        self.assertEqual(normalize_tipo_te(""), "Sem Informa\u00e7\u00e3o")


if __name__ == "__main__":
    unittest.main()
