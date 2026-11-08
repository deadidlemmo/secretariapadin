"""Conferência tripla com nomes e documentos inteiramente fictícios."""

import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from flask import Flask
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

from services.confere_diario import compare_three_sources, extract_diario_pdf_records, extract_diario_pdfs
from confere import confere_bp
from services.confere_escolas import get_confere_school_config
from services.conferir_listas import build_excel_report, normalize_name_for_match, run_conferencia


def make_diario_pdf(rows):
    output = io.BytesIO()
    story = [
        Paragraph("EM Jose Padin Mouta Ano: 2026", getSampleStyleSheet()["Normal"]),
        Paragraph("3º Trimestre Turno: Tarde Turma: 3º Ano A(I) - POLIVALENTE", getSampleStyleSheet()["Normal"]),
    ]
    table = Table(
        [["Nº", "Nome do Aluno", "Movimentação", "P", "M", "C", "H", "G", "E", "A", "F", "Total"]] +
        [[str(number), name, movement] + [""] * 9 for number, name, movement in rows],
        colWidths=[25, 185, 190] + [36] * 9,
    )
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    story.extend([table, PageBreak(), Paragraph("Assinaturas", getSampleStyleSheet()["Normal"])])
    SimpleDocTemplate(output, pagesize=landscape(A4), leftMargin=12, rightMargin=12).build(story)
    return output.getvalue()


def make_lista():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "LISTA CORRIDA"
    sheet.append(["SÉRIE", "Nº", "RM", "NOME", "SEXO", "DATA NASC.", "RA", "COD.", "OBS"])
    sheet.append(["3ºA", 7, "900001", "ALUNO FICTICIO A", "M", "01/01/2017", "990003", "MA", ""])
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def make_sed_pdf():
    output = io.BytesIO()
    table = Table(
        [
            ["Serie", "N", "Nome do Aluno", "RA", "Digito", "UF", "Nascimento", "Movimentacao", "Situacao"],
            ["3", "99", "ALUNO FICTICIO A", "990003", "", "SP", "01/01/2017", "", "ATIVO"],
        ],
        colWidths=[34, 24, 210, 66, 40, 30, 85, 85, 66],
    )
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    story = [Paragraph("Turma: 3º ANO A TARDE ANUAL", getSampleStyleSheet()["Normal"]), table]
    SimpleDocTemplate(output, pagesize=landscape(A4), leftMargin=18, rightMargin=18).build(story)
    return output.getvalue()


def record(source, name, turma, status, number="", ra="990003", other=""):
    data = {
        "source": source, "nome": name,
        "nome_match_norm": normalize_name_for_match(name),
        "turma_key": turma, "situacao": status,
        "numero_chamada": number,
        "ra_keys": [ra] if source != "diario" else [],
    }
    if source == "diario":
        data.update({"movimentacao": "", "outra_turma_key": other, "pdf_origem": "diario_ficticio.pdf"})
    return data


class DiarioConferenceTests(unittest.TestCase):
    def test_parser_recognizes_both_remanejos_and_skips_signature_page(self):
        pdf = make_diario_pdf([
            (1, "ALUNO FICTICIO A", "REM. DE 3º Ano B(I) - 09/02"),
            (2, "ALUNO FICTICIO B", "REM. P/ 3º Ano C(I) - 27/07"),
            (3, "ALUNO FICTICIO C", "T.E. - 25/09"),
        ])
        rows = extract_diario_pdf_records(pdf, "diario_ficticio.pdf")
        self.assertEqual([r["situacao"] for r in rows], ["REM_DE", "REM_P", "TE"])
        self.assertEqual([r["outra_turma_key"] for r in rows[:2]], ["3B", "3C"])
        self.assertEqual([r["numero_chamada"] for r in rows], ["1", "2", "3"])

    def test_call_number_does_not_identify_another_student(self):
        lista = [record("lista", "ALUNO FICTICIO A", "3A", "MA", "1")]
        sed = [record("sed", "ALUNO FICTICIO A", "3A", "ATIVO")]
        diario = [record("diario", "ALUNO FICTICIO B", "3A", "ATIVO", "1")]
        result = compare_three_sources(lista, sed, {"records": diario, "successful_files": ["diario_ficticio.pdf"], "errors": [], "duplicate_files": []})
        self.assertEqual(result["summary"]["total_alunos"], 2)
        self.assertTrue(all(row["pendencias"] for row in result["rows"]))

    def test_two_versions_of_same_class_are_rejected(self):
        files = [
            {"filename": "diario_antigo_ficticio.pdf", "content": make_diario_pdf([(1, "ALUNO FICTICIO A", "")])},
            {"filename": "diario_novo_ficticio.pdf", "content": make_diario_pdf([(2, "ALUNO FICTICIO B", "")])},
        ]
        with self.assertRaisesRegex(ValueError, "mais de um PDF"):
            extract_diario_pdfs(files)

    def test_remanejo_history_resolves_to_destination(self):
        name = "ALUNO FICTICIO REMANEJADO"
        lista = [record("lista", name, "3B", "MA", "4")]
        sed = [record("sed", name, "3A", "REMA"), record("sed", name, "3B", "ATIVO")]
        diario = [record("diario", name, "3A", "REM_P", "8", other="3B"),
                  record("diario", name, "3B", "REM_DE", "4", other="3A")]
        result = compare_three_sources(lista, sed, {"records": diario, "successful_files": ["diario_ficticio.pdf"], "errors": [], "duplicate_files": []})
        self.assertEqual(result["summary"]["total_ok"], 1)
        self.assertEqual(result["rows"][0]["turma_diario"], "3ºB")

    def test_upload_pipeline_and_excel_export_include_triple_result(self):
        result = run_conferencia(
            make_lista(), [{"filename": "sed_ficticio.pdf", "content": make_sed_pdf()}],
            school_config=get_confere_school_config("padin"),
            diario_file_items=[{"filename": "diario_ficticio.pdf", "content": make_diario_pdf([(7, "ALUNO FICTICIO A", "")])}],
        )
        self.assertEqual(result["diario"]["summary"]["total_ok"], 1)
        self.assertEqual(result["diario"]["summary"]["total_pendencias"], 0)
        export = load_workbook(build_excel_report(result))
        self.assertIn("Diário pendências", export.sheetnames)
        self.assertIn("Diário completo", export.sheetnames)
        self.assertEqual(export["Diário completo"].max_row, 2)

    def test_padin_form_accepts_diario_and_renders_triple_result(self):
        root = Path(__file__).resolve().parents[1]
        app = Flask(__name__, template_folder=str(root / "templates"), static_folder=str(root / "static"))
        app.config.update(TESTING=True, SECRET_KEY="fictitious-test-session-key")
        app.register_blueprint(confere_bp, url_prefix="/confere")
        with tempfile.TemporaryDirectory() as folder, patch("confere.RESULTS_FOLDER", folder):
            with app.test_client() as client:
                with client.session_transaction() as session:
                    session["confere_logged_in"] = True
                response = client.post("/confere/", data={
                    "school_id": "padin",
                    "lista_piloto": (make_lista(), "lista_ficticia.xlsx"),
                    "sed_pdfs": (io.BytesIO(make_sed_pdf()), "sed_ficticio.pdf"),
                    "diario_pdfs": (io.BytesIO(make_diario_pdf([(7, "ALUNO FICTICIO A", "")])), "diario_ficticio.pdf"),
                }, follow_redirects=True)
                self.assertEqual(response.status_code, 200)
                html = response.get_data(as_text=True)
                self.assertIn("Conferência tripla", html)
                self.assertIn("confere-diario.css?v=1", html)
                self.assertIn("ALUNO FICTICIO A", html)
                self.assertIn("Sem pendências", html)


if __name__ == "__main__":
    unittest.main()
