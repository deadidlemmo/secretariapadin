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


def make_lista(ra="990003", name="ALUNO FICTICIO A"):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "LISTA CORRIDA"
    sheet.append(["SÉRIE", "Nº", "RM", "NOME", "SEXO", "DATA NASC.", "RA", "COD.", "OBS"])
    sheet.append(["3ºA", 7, "900001", name, "M", "01/01/2017", ra, "MA", ""])
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def make_sed_pdf(name="ALUNO FICTICIO A", digit=""):
    output = io.BytesIO()
    table = Table(
        [
            ["Serie", "N", "Nome do Aluno", "RA", "Digito", "UF", "Nascimento", "Movimentacao", "Situacao"],
            ["3", "99", name, "990003", digit, "SP", "01/01/2017", "", "ATIVO"],
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

    def test_call_only_alert_is_separate_from_enrollment_alerts(self):
        name = "ALUNO FICTICIO A"
        lista = [record("lista", name, "3A", "MA", "7")]
        sed = [record("sed", name, "3A", "ATIVO")]
        diario = [record("diario", name, "3A", "ATIVO", "8")]
        result = compare_three_sources(lista, sed, {"records": diario, "successful_files": ["diario_ficticio.pdf"], "errors": [], "duplicate_files": []})
        self.assertEqual(result["summary"]["total_somente_chamada"], 1)
        self.assertEqual(result["summary"]["total_outros_alertas"], 0)
        self.assertEqual(result["rows"][0]["tipo_principal"], "chamada")

    def test_one_letter_diario_typo_is_one_name_alert(self):
        lista_name = "ALUNO FICTICIO OLIMPIO"
        diario_name = "ALUNO FICTICIO OLIMPO"
        result = run_conferencia(
            make_lista(name=lista_name),
            [{"filename": "sed_ficticio.pdf", "content": make_sed_pdf(name=lista_name, digit="0")}],
            school_config=get_confere_school_config("padin"),
            diario_file_items=[{"filename": "diario_ficticio.pdf", "content": make_diario_pdf([(7, diario_name, "")])}],
        )
        self.assertEqual(result["diario"]["summary"]["total_alunos"], 1)
        self.assertEqual(result["diario"]["summary"]["totais_tipos"]["nome"], 1)
        row = result["diario"]["rows"][0]
        self.assertEqual(row["tipo_principal"], "nome")
        self.assertEqual(row["pendencias_tipos"], ["nome"])
        self.assertEqual((row["nome_lista"], row["nome_sed"], row["nome_diario"]),
                         (lista_name, lista_name, diario_name))
        self.assertNotIn("não localizado", " ".join(row["pendencias"]).lower())
        export = load_workbook(build_excel_report(result))
        self.assertEqual(export["Diário pendências"]["B1"].value, "Nome Diário")
        self.assertEqual(export["Diário pendências"]["B2"].value, diario_name)

    def test_typo_is_not_merged_when_call_differs_or_candidates_are_ambiguous(self):
        lista_name = "ALUNO FICTICIO OLIMPIO"
        diario_name = "ALUNO FICTICIO OLIMPO"
        source = {"successful_files": ["diario_ficticio.pdf"], "errors": [], "duplicate_files": []}
        wrong_call = compare_three_sources(
            [record("lista", lista_name, "3A", "MA", "7")],
            [record("sed", lista_name, "3A", "ATIVO")],
            {**source, "records": [record("diario", diario_name, "3A", "ATIVO", "8")]},
        )
        self.assertEqual(wrong_call["summary"]["total_alunos"], 2)

        conflicting_ra = compare_three_sources(
            [record("lista", lista_name, "3A", "MA", "7", ra="111111")],
            [record("sed", lista_name, "3A", "ATIVO", ra="222222")],
            {**source, "records": [record("diario", diario_name, "3A", "ATIVO", "7")]},
        )
        self.assertEqual(conflicting_ra["summary"]["total_alunos"], 2)

        anchors = ["ALUNO FICTICIO MARTA", "ALUNO FICTICIO MARIA"]
        lista = [record("lista", name, "3A", "MA", "7", ra=str(index)) for index, name in enumerate(anchors, 1)]
        sed = [record("sed", name, "3A", "ATIVO", ra=str(index)) for index, name in enumerate(anchors, 1)]
        ambiguous = compare_three_sources(
            lista, sed, {**source, "records": [record("diario", "ALUNO FICTICIO MARLA", "3A", "ATIVO", "7")]},
        )
        self.assertEqual(ambiguous["summary"]["total_alunos"], 3)

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
        self.assertEqual(export["Diário completo"]["F1"].value, "RA Lista")
        self.assertEqual(export["Diário completo"]["G1"].value, "RA SED")
        self.assertEqual(export["Diário completo"]["F2"].value, result["diario"]["rows"][0]["ra_lista"])

    def test_zero_ra_is_reported_as_missing_in_both_views(self):
        result = run_conferencia(
            make_lista(ra="0"), [{"filename": "sed_ficticio.pdf", "content": make_sed_pdf()}],
            school_config=get_confere_school_config("padin"),
            diario_file_items=[{"filename": "diario_ficticio.pdf", "content": make_diario_pdf([(7, "ALUNO FICTICIO A", "")])}],
        )
        self.assertIn("RA não informado na Lista Piloto.", result["rows"][0]["observacao"])
        self.assertIn("RA não informado na Lista Piloto", result["diario"]["rows"][0]["pendencias"][0])
        self.assertEqual(result["diario"]["rows"][0]["tipo_principal"], "identidade")

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
                self.assertTrue("Lista Piloto × SED × Diário Online" in html, "Título da conferência tripla ausente")
                self.assertIn("confere-diario.css?v=5", html)
                self.assertIn("ALUNO FICTICIO A", html)
                self.assertIn('data-triple-filter="dados"', html)
                self.assertIn('data-triple-filter="nome"', html)
                self.assertIn('data-triple-status="ok"', html)
                self.assertIn('id="diario-result-search"', html)

                lista_name = "ALUNO FICTICIO OLIMPIO"
                diario_name = "ALUNO FICTICIO OLIMPO"
                typo_response = client.post("/confere/", data={
                    "school_id": "padin",
                    "lista_piloto": (make_lista(name=lista_name), "lista_ficticia.xlsx"),
                    "sed_pdfs": (io.BytesIO(make_sed_pdf(name=lista_name)), "sed_ficticio.pdf"),
                    "diario_pdfs": (io.BytesIO(make_diario_pdf([(7, diario_name, "")])), "diario_ficticio.pdf"),
                }, follow_redirects=True)
                self.assertEqual(typo_response.status_code, 200)
                typo_html = typo_response.get_data(as_text=True)
                self.assertEqual(typo_html.count('data-triple-types="nome"'), 1)
                self.assertIn("Nome cadastrado de forma diferente no Diário Online", typo_html)
                self.assertRegex(typo_html, r'Lista Piloto</span>\s*<strong class="comparison-name">ALUNO FICTICIO OLIMPIO</strong>')
                self.assertRegex(typo_html, r'Diário Online</span>\s*<strong class="comparison-name">ALUNO FICTICIO OLIMPO</strong>')
                self.assertNotIn("Aluno ativo não localizado", typo_html)


if __name__ == "__main__":
    unittest.main()
