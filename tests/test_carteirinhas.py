import tempfile
import unittest
from pathlib import Path

import pandas as pd

from services.carteirinhas import (
    build_carteirinhas_context,
    mark_carteirinhas_impressas,
    normalize_rms,
)


def write_lista_corrida(path):
    lista = pd.DataFrame(
        [
            {
                "RM": 101,
                "NOME": "Aluno Com Foto",
                "DATA NASC.": "2020-01-02",
                "RA": "RA101",
                "SAI SOZINHO?": "SIM",
                "S\u00c9RIE": "3\u00baA",
                "HOR\u00c1RIO": "Segunda \u00e0 Sexta-Feira das 13h00 \u00e0s 17h00",
            },
            {
                "RM": 102,
                "NOME": "Aluno Sem Foto",
                "DATA NASC.": "invalida",
                "RA": "RA102",
                "SAI SOZINHO?": "NAO",
                "S\u00c9RIE": "4\u00baA",
                "HOR\u00c1RIO": "08h00",
            },
            {
                "RM": 0,
                "NOME": "Aluno Ignorado",
                "DATA NASC.": "2020-01-02",
                "RA": "RA000",
                "SAI SOZINHO?": "SIM",
                "S\u00c9RIE": "5\u00baA",
                "HOR\u00c1RIO": "08h00",
            },
        ]
    )
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        lista.to_excel(writer, sheet_name="LISTA CORRIDA", index=False)


class CarteirinhasServiceTests(unittest.TestCase):
    def test_normalize_rms_keeps_unique_positive_ints(self):
        self.assertEqual(normalize_rms(["101", "0", "abc", 101, "102"]), [101, 102])

    def test_build_carteirinhas_context_filters_and_paginates(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir, "lista.xlsx")
            write_lista_corrida(path)

            context = build_carteirinhas_context(
                path,
                ano=2026,
                get_printed_set_func=lambda ano: {102},
                get_photo_url_func=lambda rm: f"/static/fotos/{rm}.jpg" if rm == 101 else None,
            )

        alunos = context["pages"][0]
        self.assertEqual(context["ano"], 2026)
        self.assertEqual(len(alunos), 2)
        self.assertEqual(alunos[0]["rm"], 101)
        self.assertEqual(alunos[0]["data_nasc"], "02/01/2020")
        self.assertEqual(alunos[0]["horario_dias"], "Segunda \u00e0 Sexta-Feira")
        self.assertEqual(alunos[0]["horario_faixa"], "13h00 \u00e0s 17h00")
        self.assertEqual(alunos[0]["classe_cor"], "verde")
        self.assertEqual(alunos[0]["foto_url"], "/static/fotos/101.jpg")
        self.assertFalse(alunos[0]["impresso"])
        self.assertEqual(alunos[1]["status_texto"], "N\u00e3o sai sozinho")
        self.assertEqual(alunos[1]["horario_dias"], "")
        self.assertEqual(alunos[1]["horario_faixa"], "08h00")
        self.assertTrue(alunos[1]["impresso"])
        self.assertEqual(context["total_sem_foto"], 1)
        self.assertEqual(context["alunos_sem_foto"][0]["rm"], 102)

    def test_build_carteirinhas_context_applies_filters(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir, "lista.xlsx")
            write_lista_corrida(path)

            only_photo = build_carteirinhas_context(
                path,
                somente_com_foto=True,
                ano=2026,
                get_printed_set_func=lambda ano: {102},
                get_photo_url_func=lambda rm: f"/static/fotos/{rm}.jpg" if rm == 101 else None,
            )
            only_not_printed = build_carteirinhas_context(
                path,
                somente_nao_impressas=True,
                ano=2026,
                get_printed_set_func=lambda ano: {102},
                get_photo_url_func=lambda rm: f"/static/fotos/{rm}.jpg" if rm == 101 else None,
            )

        self.assertEqual([aluno["rm"] for aluno in only_photo["pages"][0]], [101])
        self.assertEqual([aluno["rm"] for aluno in only_not_printed["pages"][0]], [101])

    def test_mark_carteirinhas_impressas_filters_by_photo(self):
        calls = {}

        def fake_mark(ano, rms):
            calls["ano"] = ano
            calls["rms"] = rms
            return 1, 2

        result = mark_carteirinhas_impressas(
            ["101", "102", "102", "abc"],
            ano=2026,
            student_has_photo_func=lambda rm: rm == 101,
            mark_printed_rms_func=fake_mark,
        )

        self.assertEqual(calls, {"ano": 2026, "rms": [101]})
        self.assertEqual(
            result,
            {
                "ok": True,
                "ano": 2026,
                "received": 2,
                "considered_with_photo": 1,
                "added": 1,
                "total_printed": 2,
            },
        )


if __name__ == "__main__":
    unittest.main()
