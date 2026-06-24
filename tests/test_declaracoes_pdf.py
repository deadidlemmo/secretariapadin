import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import app as app_module
import pandas as pd
from services import declaracoes_pdf

app = app_module.app


class DeclaracoesPdfTests(unittest.TestCase):
    def _write_lista_fundamental(self, path):
        df = pd.DataFrame(
            [
                {
                    "RM": 19034,
                    "NOME": "ALUNO FICTICIO",
                    "SÉRIE": "2ºA",
                    "DATA NASC.": "10/03/2018",
                    "RA": "123456789X",
                    "HORÁRIO": "13h30 às 17h30",
                }
            ]
        )
        with pd.ExcelWriter(path) as writer:
            df.to_excel(writer, sheet_name="LISTA CORRIDA", index=False)

    def test_prepare_html_for_pdf_uses_static_file_uri(self):
        html = '<img src="/static/logos/escola.png"><link href="/static/css/app.css">'
        prepared = declaracoes_pdf.prepare_html_for_pdf(html, str(Path.cwd()))

        self.assertIn((Path.cwd() / "static").resolve().as_uri(), prepared)
        self.assertNotIn('src="/static/', prepared)
        self.assertNotIn('href="/static/', prepared)

    def test_render_declaracao_pdf_bytes_uses_browser_without_header_footer(self):
        old_dir = os.environ.get("DECLARACAO_PDF_DIR")
        old_browser = os.environ.get("DECLARACAO_PDF_BROWSER_PATH")

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            fake_browser = tmp_path / "chrome.exe"
            fake_browser.write_text("", encoding="utf-8")
            os.environ["DECLARACAO_PDF_DIR"] = str(tmp_path)
            os.environ["DECLARACAO_PDF_BROWSER_PATH"] = str(fake_browser)

            token = declaracoes_pdf.create_declaracao_pdf_token(
                "<html><body>Declaracao</body></html>",
                str(Path.cwd()),
            )
            captured = {}

            def fake_run(cmd, **kwargs):
                captured["cmd"] = cmd
                pdf_arg = next(arg for arg in cmd if arg.startswith("--print-to-pdf="))
                Path(pdf_arg.split("=", 1)[1]).write_bytes(b"%PDF-1.4\n%%EOF")
                return SimpleNamespace(returncode=0, stdout="", stderr="")

            try:
                with patch.object(declaracoes_pdf.subprocess, "run", side_effect=fake_run):
                    pdf_bytes = declaracoes_pdf.render_declaracao_pdf_bytes(token)
            finally:
                if old_dir is None:
                    os.environ.pop("DECLARACAO_PDF_DIR", None)
                else:
                    os.environ["DECLARACAO_PDF_DIR"] = old_dir
                if old_browser is None:
                    os.environ.pop("DECLARACAO_PDF_BROWSER_PATH", None)
                else:
                    os.environ["DECLARACAO_PDF_BROWSER_PATH"] = old_browser

        self.assertEqual(pdf_bytes, b"%PDF-1.4\n%%EOF")
        self.assertIn("--headless=new", captured["cmd"])
        self.assertIn("--no-pdf-header-footer", captured["cmd"])
        self.assertTrue(any(arg.startswith("--print-to-pdf=") for arg in captured["cmd"]))

    def test_declaracao_pdf_route_returns_pdf_response(self):
        with app.test_client() as client:
            with client.session_transaction() as flask_session:
                flask_session["logged_in"] = True

            with patch("app.render_declaracao_pdf_bytes", return_value=b"%PDF-1.4\n%%EOF"):
                response = client.get("/declaracao/pdf/0123456789abcdef0123456789abcdef")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/pdf")
        self.assertEqual(response.data, b"%PDF-1.4\n%%EOF")

    def test_render_declaracao_pdf_response_returns_pdf_directly(self):
        with app.test_request_context("/declaracao/tipo", method="POST"):
            with patch("app.create_declaracao_pdf_token", return_value="abc"), patch(
                "app.render_declaracao_pdf_bytes",
                return_value=b"%PDF-1.4\n%%EOF",
            ):
                response = app_module.render_declaracao_pdf_response(
                    titulo="Declaração",
                    data_extenso="Praia Grande, 18 de junho de 2026",
                    declaracao_text="<p>Conteúdo fictício.</p>",
                    additional_css="",
                    body_classes=[],
                    print_body_padding="0.5cm 0.5cm",
                )

        self.assertEqual(response.mimetype, "application/pdf")
        self.assertIn("inline", response.headers.get("Content-Disposition", ""))

    def test_declaracao_tipo_post_returns_html_preview_by_default(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            lista_path = Path(tmpdir) / "lista_ficticia.xlsx"
            self._write_lista_fundamental(lista_path)

            with app.test_client() as client:
                with client.session_transaction() as flask_session:
                    flask_session["logged_in"] = True
                    flask_session["lista_fundamental"] = str(lista_path)

                with patch(
                    "app.render_declaracao_pdf_bytes",
                    return_value=b"%PDF-1.4\n%%EOF",
                ):
                    response = client.post(
                        "/declaracao/tipo",
                        data={
                            "segmento_escolhido": "Fundamental",
                            "rm": "19034",
                            "tipo": "Escolaridade",
                        },
                    )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "text/html")
        self.assertIn(b"<html", response.data.lower())
        self.assertIn(b"Imprimir", response.data)

    def test_declaracao_tipo_post_can_still_return_pdf_when_requested(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            lista_path = Path(tmpdir) / "lista_ficticia.xlsx"
            self._write_lista_fundamental(lista_path)

            with app.test_client() as client:
                with client.session_transaction() as flask_session:
                    flask_session["logged_in"] = True
                    flask_session["lista_fundamental"] = str(lista_path)

                with patch(
                    "app.render_declaracao_pdf_bytes",
                    return_value=b"%PDF-1.4\n%%EOF",
                ):
                    response = client.post(
                        "/declaracao/tipo",
                        data={
                            "segmento_escolhido": "Fundamental",
                            "rm": "19034",
                            "tipo": "Escolaridade",
                            "preview": "0",
                        },
                    )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/pdf")
        self.assertNotIn(b"<html", response.data.lower())


if __name__ == "__main__":
    unittest.main()
