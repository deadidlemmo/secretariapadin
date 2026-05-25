import ast
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
            if name in {"upload_excel", "index"}:
                self.assertTrue(any("confere_login_required" == decorator for decorator in decorators))
                protected.add(name)
        self.assertEqual(protected, {"upload_excel", "index"})

    def test_holidays_path_uses_existing_file(self):
        source = read_text("app.py") + read_text("config.py")
        self.assertNotIn("feriados_nacionais.json", source)
        self.assertIn("feriados.json", source)
        self.assertTrue((ROOT / "modelos" / "feriados.json").exists())

    def test_gitignore_protects_sensitive_runtime_paths(self):
        gitignore = read_text(".gitignore")
        for expected in ["uploads/", "static/fotos/", "__pycache__/", "*.py[cod]", ".env", "env/", "venv/"]:
            self.assertIn(expected, gitignore)

    def test_declaracao_tipo_assets_are_externalized(self):
        template = read_text("templates/declaracao_tipo.html")
        js = read_text("static/js/declaracao_tipo.js")

        self.assertNotIn("{% block extra_styles %}", template)
        self.assertNotIn("  <script>\n", template)
        self.assertIn("css/declaracao_tipo.css", template)
        self.assertIn("js/declaracao_tipo.js", template)
        self.assertTrue((ROOT / "static" / "css" / "declaracao_tipo.css").exists())
        self.assertTrue((ROOT / "static" / "js" / "declaracao_tipo.js").exists())
        self.assertNotIn("{{", js)
        self.assertNotIn("{%", js)

    def test_carteirinhas_assets_are_externalized(self):
        template = read_text("templates/gerar_carteirinhas.html")
        js = read_text("static/js/carteirinhas.js")

        self.assertNotIn("<style>", template)
        self.assertNotIn("<script>\n", template)
        self.assertIn("css/carteirinhas.css", template)
        self.assertIn("js/carteirinhas.js", template)
        self.assertTrue((ROOT / "static" / "css" / "carteirinhas.css").exists())
        self.assertTrue((ROOT / "static" / "js" / "carteirinhas.js").exists())
        self.assertNotIn("{{", js)
        self.assertNotIn("{%", js)


if __name__ == "__main__":
    unittest.main()
