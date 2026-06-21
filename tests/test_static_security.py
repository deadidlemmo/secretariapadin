import ast
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read_text(path):
    return (ROOT / path).read_text(encoding="utf-8")


def parse(path):
    return ast.parse(read_text(path))


def route_functions(path):
    tree = parse(path)
    for node in tree.body:
        if not isinstance(node, ast.FunctionDef):
            continue
        decorators = [ast.unparse(d) for d in node.decorator_list]
        routes = [d for d in decorators if ".route(" in d]
        if routes:
            yield node.name, decorators, routes


HEX_COLOR_RE = re.compile(r"#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})\b")
CSS_RULE_RE = re.compile(r"([^{}]+)\{([^{}]+)\}", re.S)
CSS_PROP_RE = re.compile(r"([\w-]+)\s*:\s*([^;]+);")


def expand_hex_color(value):
    value = value.lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def relative_luminance(rgb):
    channels = []
    for channel in rgb:
        channel = channel / 255
        if channel <= 0.03928:
            channels.append(channel / 12.92)
        else:
            channels.append(((channel + 0.055) / 1.055) ** 2.4)
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def contrast_ratio(foreground, background):
    first = relative_luminance(foreground)
    second = relative_luminance(background)
    return (max(first, second) + 0.05) / (min(first, second) + 0.05)


class StaticSecurityTests(unittest.TestCase):
    def test_no_default_secrets_in_source(self):
        combined = read_text("app.py") + read_text("confere.py") + read_text("config.py")
        self.assertNotIn("sua_chave_secreta", combined)
        self.assertNotIn("minha_senha", combined)
        self.assertNotIn('ACCESS_TOKEN = "change-me"', combined)

    def test_quantinclusao_requires_login(self):
        routes = dict((name, decorators) for name, decorators, _routes in route_functions("app.py"))
        self.assertIn("quantinclusao", routes)
        self.assertTrue(any("login_required" == decorator for decorator in routes["quantinclusao"]))

    def test_confere_routes_require_login(self):
        protected = set()
        for name, decorators, _routes in route_functions("confere.py"):
            if name in {"index", "resultado", "exportar_excel"}:
                self.assertTrue(any("confere_login_required" == decorator for decorator in decorators))
                protected.add(name)
        self.assertEqual(protected, {"index", "resultado", "exportar_excel"})

    def test_holidays_path_uses_existing_file(self):
        source = read_text("app.py") + read_text("config.py")
        self.assertNotIn("feriados_nacionais.json", source)
        self.assertIn("feriados.json", source)
        self.assertTrue((ROOT / "modelos" / "feriados.json").exists())

    def test_gitignore_protects_sensitive_runtime_paths(self):
        gitignore = read_text(".gitignore")
        for expected in ["uploads/", "__pycache__/", "*.py[cod]", ".env", "env/", "venv/"]:
            self.assertIn(expected, gitignore)
        self.assertNotIn("static/fotos/", gitignore)

    def test_declaracao_tipo_assets_are_externalized(self):
        template = read_text("templates/declaracao_tipo.html")
        css = read_text("static/css/declaracao_tipo.css")
        js = read_text("static/js/declaracao_tipo.js")

        self.assertNotIn("{% block extra_styles %}", template)
        self.assertNotIn("  <script>\n", template)
        self.assertIn("css/declaracao_tipo.css", template)
        self.assertIn("js/declaracao_tipo.js", template)
        self.assertTrue((ROOT / "static" / "css" / "declaracao_tipo.css").exists())
        self.assertTrue((ROOT / "static" / "js" / "declaracao_tipo.js").exists())
        self.assertIn("var(--app-control-radius)", css)
        self.assertIn("var(--app-pill-radius)", css)
        self.assertIn("#declaracao-page .btn-declaracao.step-highlight-field", css)
        self.assertIn("#declaracao-page .btn-lote.step-highlight-field", css)
        self.assertIn("background: linear-gradient(135deg, var(--app-primary-dark), var(--app-primary)) !important", css)
        self.assertIn("#declaracao-page .btn-declaracao.step-highlight-field:disabled", css)
        self.assertIn("#declaracao-page .btn-declaracao:disabled i", css)
        self.assertIn("liberarTelaAposAbrirPdf", js)
        self.assertIn("window.setTimeout(ocultarOverlay, 1200)", js)
        self.assertNotIn("{{", js)
        self.assertNotIn("{%", js)

    def test_docker_deploy_files_preserve_pdf_and_school_assets(self):
        dockerfile = read_text("Dockerfile")
        dockerignore = read_text(".dockerignore")
        env_example = read_text(".env.example")
        readme = read_text("README.md")
        check_pdf = read_text("scripts/check_pdf_environment.py")

        self.assertIn("FROM python:3.12-slim-bookworm", dockerfile)
        self.assertIn("chromium", dockerfile)
        self.assertIn("DECLARACAO_PDF_BROWSER_PATH=/usr/bin/chromium", dockerfile)
        self.assertIn("gunicorn app:app", dockerfile)
        self.assertIn("WEB_CONCURRENCY:-3", dockerfile)
        self.assertIn("DECLARACAO_PDF_BROWSER_PATH=", env_example)
        self.assertIn("Render com PDF de Declaracoes", readme)
        self.assertIn("scripts/check_pdf_environment.py", readme)
        self.assertIn("find_browser_executable", check_pdf)
        self.assertIn("render_declaracao_pdf_bytes", check_pdf)

        self.assertIn("uploads/*", dockerignore)
        self.assertIn("!uploads/escolas.csv", dockerignore)
        self.assertNotIn("static/fotos", dockerignore)
        self.assertNotIn("modelos/", dockerignore)

    def test_declaracao_tipo_preserves_form_contracts(self):
        template = read_text("templates/declaracao_tipo.html")

        for endpoint in ["declaracao_tipo", "escolas_search"]:
            self.assertIn(f"url_for('{endpoint}'", template)
        for expected in [
            'id="declaracao-page"',
            'data-segmento=',
            'data-tem-lista=',
            'id="form-declaracao"',
            'id="form-declaracao-personalizada"',
            'name="segmento_escolhido"',
            'id="rm"',
            'name="rm"',
            'id="tipo"',
            'name="tipo"',
            'name="excel_file"',
            'name="modo_declaracao"',
            'name="tipo_declaracao_personalizada"',
            'name="nome_aluno"',
            "declaracao-compact-label",
            "declaracao-full-select",
            "frequencia-help-text",
            "frequencia-footnote",
            'formtarget="_blank"',
            "Gerar declaração em PDF",
            "Gerar declaração personalizada em PDF",
        ]:
            self.assertIn(expected, template)
        self.assertIn("declaracao-page-header", template)
        self.assertIn("segmento-card-arrow", template)
        self.assertIn("declaracao-upload-zone", template)
        self.assertIn("status-alert-content", template)

    def test_declaracao_print_supports_siae_model_layout(self):
        template = read_text("templates/declaracao_print.html")
        service = read_text("services/declaracoes.py")

        for expected in [
            "modelo-siae",
            "siae-declaration",
            "siae-option",
            "siae-check",
            "siae-date",
            "declaracao-render",
            "siae-student-line",
            "siae-option-detail",
            "siae-closing",
            "siae-observations",
            "siae-observations-content",
            "is_modelo_siae",
            "pdf_url",
            "acoes-preview",
            "Gerar PDF da Declaração",
            "max-width: 182mm",
            "min-height: calc(297mm - 20mm)",
            "margin: 10mm 14mm",
            "font-family: Arial, sans-serif",
            "font-size: 13.5pt",
            "font-size: 18pt",
            "font-size: 12pt",
            "border-bottom: 0",
            "border-top: 1px solid #777",
            "border-bottom: 1px solid #bbb",
            "border-radius: 0",
            "grid-template-columns: 32px minmax(0, 1fr)",
            "align-items: center",
            "place-items: center",
            "align-self: center",
            "text-decoration-line: line-through",
            "grid-template-rows: auto minmax(0, 1fr) auto",
            "margin-top: auto",
            "font-size: 13.5pt",
            "font-weight: 400",
        ]:
            self.assertIn(expected, template)

        for expected in [
            "_build_modelo_siae_declaracao_html",
            "siae-option-{key}",
            '_build_siae_option("conclusao"',
            '_build_siae_option("escolaridade"',
            '_build_siae_option("transferencia"',
            "_build_siae_observations_html",
            "_build_historico_unidade_observacao",
            "bloco-observacoes",
            "paragrafo-observacao",
            "conteudo-observacoes",
            "N\\u00e3o preenchido",
            "Selecionada",
        ]:
            self.assertIn(expected, service)

        for removed in [
            "<ul>",
            "<li>",
        ]:
            self.assertNotIn(removed, service)

        for removed in [
            "siae-modelo-label",
            "siae-municipal-heading",
            "siae-option-status",
            "Georgia",
            "Times New Roman",
            "window.print()",
        ]:
            self.assertNotIn(removed, template)

        for removed in [
            "siae-modelo-label",
            "siae-municipal-heading",
            "na cidade de",
            "Estado de <span",
        ]:
            self.assertNotIn(removed, service)

    def test_carteirinhas_assets_are_externalized(self):
        template = read_text("templates/gerar_carteirinhas.html")
        css = read_text("static/css/carteirinhas.css")
        js = read_text("static/js/carteirinhas.js")

        self.assertNotIn("<style>", template)
        self.assertNotIn("<script>\n", template)
        self.assertIn("css/carteirinhas.css", template)
        self.assertIn("js/carteirinhas.js", template)
        self.assertTrue((ROOT / "static" / "css" / "carteirinhas.css").exists())
        self.assertTrue((ROOT / "static" / "js" / "carteirinhas.js").exists())
        self.assertTrue((ROOT / "static" / "logos" / "escola.png").exists())
        self.assertNotIn("{{", js)
        self.assertNotIn("{%", js)
        self.assertIn("logos/escola.png", template)
        self.assertNotIn(">JP<", template)
        self.assertNotIn("filtro-foto-pill", template)
        for expected in [
            "carteirinhas-screen-header",
            "carteirinhas-toolbar",
            "toolbar-actions-row",
            "loading-overlay-content",
            "card-school-mark",
            "card-school-logo",
            "card-year-top",
            "photo-shell",
            "foto-placeholder",
            "horario-row",
            "horario-dia",
            "horario-faixa",
            "status-symbol",
            "status-text",
            "carteirinhas-footer",
            "filter-submit-noscript",
            "no-cards-message",
        ]:
            self.assertIn(expected, template)
        for expected in [
            ".card-school-name",
            ".card-document-title",
            ".card-year-top",
            ".photo-caption",
            ".horario-value",
            ".horario-dia",
            ".horario-faixa",
            ".status-text",
            ".loading-overlay-content",
            ".filter-submit-noscript",
            ".no-cards-message",
            ".upload-success-toast",
            ".carteirinhas-screen-header",
            ".carteirinhas-footer",
            "max-width: none",
            "rgba(0, 0, 0, 0.76)",
            "#244e47",
            "width: 10cm",
            "height: 8.45cm",
            "background: #087a55",
            "background: #c5161d",
        ]:
            self.assertIn(expected, css)
        for expected in [
            "normalizarBuscaCarteirinha",
            "localizarScrollTimeout",
            "paginaTemResultado",
            "page.style.display",
            "scrollIntoView",
        ]:
            self.assertIn(expected, js)
        self.assertNotIn(".filtro-foto-pill", css)
        self.assertNotIn(">Saída<", template)
        for forbidden in ["qrcode", "qr-code", "qr_code"]:
            self.assertNotIn(forbidden, template.lower())
            self.assertNotIn(forbidden, css.lower())

    def test_global_theme_is_loaded(self):
        base = read_text("templates/base.html")
        confere = read_text("templates/index.html")
        carteirinhas = read_text("templates/gerar_carteirinhas.html")
        theme_js = read_text("static/js/theme.js")

        self.assertIn("css/app_theme.css", base)
        self.assertIn("js/theme.js", base)
        self.assertIn("js/theme.js", carteirinhas)
        self.assertIn("data-theme-toggle", base)
        self.assertIn("data-theme-toggle", carteirinhas)
        self.assertIn('{% extends "base.html" %}', confere)
        self.assertIn("confere-page", confere)
        self.assertIn('name="lista_piloto"', confere)
        self.assertIn('name="sed_pdfs"', confere)
        self.assertIn("Conferindo listas", confere)
        self.assertTrue((ROOT / "static" / "css" / "app_theme.css").exists())
        self.assertTrue((ROOT / "static" / "js" / "theme.js").exists())
        self.assertNotIn("{{", theme_js)
        self.assertNotIn("{%", theme_js)

    def test_login_uses_global_header_layout(self):
        base = read_text("templates/base.html")
        template = read_text("templates/login.html")

        self.assertIn('id="app-header"', base)
        self.assertIn("Sistema da Secretaria Escolar", base)
        self.assertIn("E.M. José Padin Mouta", base)
        self.assertIn("auth-page", template)
        self.assertIn("auth-card", template)
        self.assertNotIn("auth-header", template)

    def test_upload_listas_preserves_file_fields(self):
        template = read_text("templates/upload_listas.html")
        theme_css = read_text("static/css/app_theme.css")

        self.assertIn('id="lista_fundamental"', template)
        self.assertIn('name="lista_fundamental"', template)
        self.assertIn("required", template)
        self.assertIn('id="lista_eja"', template)
        self.assertIn('name="lista_eja"', template)
        self.assertIn("upload-drop-zone", template)
        self.assertIn('data-input-id="lista_fundamental"', template)
        self.assertIn('data-input-id="lista_eja"', template)
        self.assertIn("upload-step-copy", template)
        for expected in [
            "border-radius: 16px",
            "upload-browse",
            "var(--app-pill-radius)",
        ]:
            self.assertIn(expected, template)
        for expected in [
            "--app-control-radius: 14px",
            'input[type="file"]::file-selector-button',
            "border-radius: var(--app-pill-radius)",
        ]:
            self.assertIn(expected, theme_css)

    def test_global_form_controls_are_standardized(self):
        theme_css = read_text("static/css/app_theme.css")
        declaracao_css = read_text("static/css/declaracao_tipo.css")
        carteirinhas_css = read_text("static/css/carteirinhas.css")

        for expected in [
            "--app-control-height: 48px",
            "--app-control-padding-x: 14px",
            "--app-focus-ring",
            ".form-group",
            ".form-control:disabled",
            ".form-check-input",
            ".select2-container--default .select2-selection--single .select2-selection__rendered",
            ".form-control-file",
        ]:
            self.assertIn(expected, theme_css)

        for expected in [
            "min-height: var(--app-control-height)",
            "box-shadow: var(--app-control-shadow)",
            "height: var(--app-control-height)",
        ]:
            self.assertIn(expected, declaracao_css)

        for expected in [
            "--card-screen-control-radius: 14px",
            "--card-screen-control-height: 46px",
            "#localizarAluno::placeholder",
            "border-radius: var(--card-screen-control-radius)",
            "@media (min-width: 1024px)",
            "position: relative",
            "flex-direction: row",
            "display: contents",
            "order: 4",
            "order: 3",
            "width: min(100%, 1280px)",
            "width: min(100%, 360px)",
            "min-width: 150px",
            ".status-symbol::after",
            "border-width: 0 0.045cm 0.045cm 0",
            "transform: translate(-50%, -58%) rotate(45deg)",
            "transform: translate(-50%, -68%)",
            "transform: translate(-50%, 105%)",
        ]:
            self.assertIn(expected, carteirinhas_css)

    def test_dark_theme_has_legible_deadline_alerts(self):
        css = read_text("static/css/app_theme.css")

        for expected in [
            'html[data-theme="dark"] .deadline-card',
            'html[data-theme="dark"] .deadline-title',
            'html[data-theme="dark"] .deadline-item-title',
            'html[data-theme="dark"] .deadline-item-msg',
            'html[data-theme="dark"] .deadline-pill',
            'html[data-theme="dark"] .section-title',
            'html[data-theme="dark"] .dropzone-title',
            'html[data-theme="dark"] .field-msg.ok',
            'html[data-theme="dark"] .field-msg.error',
            'html[data-theme="dark"] .multi-prof-alert .desc',
            'html[data-theme="dark"] .missing-info-detail',
            'html[data-theme="dark"] .plan-wo-inc-students',
            'html[data-theme="dark"] select option',
            'html[data-theme="dark"] select option:checked',
        ]:
            self.assertIn(expected, css)

    def test_explicit_text_background_pairs_keep_readable_contrast(self):
        paths = list((ROOT / "static").rglob("*.css"))
        paths.extend((ROOT / "templates").rglob("*.html"))
        failures = []

        for path in paths:
            source = path.read_text(encoding="utf-8", errors="ignore")
            for selector, body in CSS_RULE_RE.findall(source):
                props = {
                    match.group(1).lower(): match.group(2).strip()
                    for match in CSS_PROP_RE.finditer(body)
                }
                foreground = props.get("color")
                background = props.get("background-color") or props.get("background")
                if not foreground or not background:
                    continue

                foreground_colors = HEX_COLOR_RE.findall(foreground)
                background_colors = HEX_COLOR_RE.findall(background)
                if not foreground_colors or not background_colors:
                    continue

                foreground_rgb = expand_hex_color(foreground_colors[-1])
                worst_ratio = min(
                    contrast_ratio(foreground_rgb, expand_hex_color(color))
                    for color in background_colors
                )
                if worst_ratio < 4.5:
                    line = source[:source.find(selector)].count("\n") + 1
                    failures.append(
                        f"{path.relative_to(ROOT)}:{line} "
                        f"{worst_ratio:.2f} {foreground_colors[-1]} on "
                        f"{', '.join(background_colors)} {selector.strip()[:80]}"
                    )

        self.assertEqual([], failures)

    def test_confere_tables_and_uploads_follow_global_theme(self):
        css = read_text("static/css/app_theme.css")

        for expected in [
            ".confere-page .table > :not(caption) > * > *",
            "--bs-table-striped-bg: var(--app-surface-muted)",
            ".confere-page .form-control[type=\"file\"]",
            ".confere-page #loadingOverlay",
            ".confere-page footer",
            ".app-loading-overlay",
            ".app-loading-panel",
            ".app-loading-spinner",
            ".app-loading-message",
        ]:
            self.assertIn(expected, css)

    def test_confere_internal_batch_upload_contract(self):
        template = read_text("templates/index.html")
        dashboard = read_text("templates/dashboard.html")

        for expected in [
            'id="confere-form"',
            'action="{{ url_for(\'confere.index\') }}"',
            'name="lista_piloto"',
            'name="sed_pdfs"',
            "excel-dropzone",
            "excel-file-count",
            "sed-scope-preview",
            "sed-scope-turmas",
            "confere.preview_pdfs",
            'accept=".pdf,application/pdf"',
            "multiple",
            "sed-dropzone",
            "sed-file-list",
            "sed-file-count",
            "Executar conferência",
            "Exportar Excel",
            "data-confere-filter",
            'data-confere-filter="pendencias"',
            "data-turma-filter",
            "data-turma-key",
            "turma-filter-row",
            "result-controls",
            "result-list",
            "data-result-group",
            "result-group-header",
            "result-card",
            "comparison-grid",
            "Nova conferência",
            "confere-result-search",
            "confere-empty-filter",
            "applyResultFilter",
            "activeTurmaFilter",
            "app-loading-overlay",
            'aria-hidden="true" style="display: none;"',
        ]:
            self.assertIn(expected, template)
        self.assertIn("url_for('confere.index')", dashboard)
        self.assertNotIn("conferealunos.onrender.com", dashboard + template)

    def test_quadro_loading_overlays_use_theme_classes(self):
        for path in [
            "templates/quadro_atendimento_mensal.html",
            "templates/quadro_quantitativo_mensal.html",
            "templates/quadros_inclusao.html",
            "templates/quantinclusao.html",
        ]:
            template = read_text(path)
            self.assertIn("app-loading-overlay", template)
            self.assertIn("app-loading-panel", template)
            self.assertNotIn("overlay.style.position", template)
            self.assertNotIn("overlay.style.background", template)

    def test_deadline_partial_uses_theme_classes(self):
        template = read_text("templates/prazos_alertas.html")
        css = read_text("static/css/app_theme.css")

        self.assertNotIn('style="', template)
        for expected in [
            "deadline-inline-panel",
            "deadline-inline-item",
            "deadline-inline-button",
        ]:
            self.assertIn(expected, template)
            self.assertIn(f".{expected}", css)

    def test_dashboard_preserves_navigation_links(self):
        template = read_text("templates/dashboard.html")

        for endpoint in ["declaracao_tipo", "carteirinhas", "quadros", "logout_route", "confere.index"]:
            self.assertIn(f"url_for('{endpoint}')", template)
        self.assertNotIn("https://conferealunos.onrender.com/", template)
        self.assertIn("dashboard-page", template)
        self.assertIn("option-card", template)
        self.assertIn("var(--app-control-radius)", template)
        self.assertIn("var(--app-pill-radius)", template)

    def test_quadro_transferencias_assets_are_externalized(self):
        template = read_text("templates/quadro_transferencias.html")
        css = read_text("static/css/quadro_transferencias.css")
        js = read_text("static/js/quadro_transferencias.js")

        self.assertNotIn("{% block extra_styles %}", template)
        self.assertNotIn("<style>", template)
        self.assertNotIn("<script>\n", template)
        self.assertNotIn("style=", template)
        self.assertIn("css/quadro_transferencias.css", template)
        self.assertIn("js/quadro_transferencias.js", template)
        self.assertTrue((ROOT / "static" / "css" / "quadro_transferencias.css").exists())
        self.assertTrue((ROOT / "static" / "js" / "quadro_transferencias.js").exists())
        self.assertIn("var(--app-control-radius", css)
        self.assertIn("var(--app-upload-gradient", css)
        self.assertIn("var(--app-pill-radius", css)
        self.assertNotIn("{{", js)
        self.assertNotIn("{%", js)


if __name__ == "__main__":
    unittest.main()
