import os
import re
import subprocess
import tempfile
import time
from pathlib import Path
from uuid import uuid4


PDF_TOKEN_RE = re.compile(r"^[a-f0-9]{32}$")
PDF_TTL_SECONDS = 60 * 60


class DeclaracaoPdfError(RuntimeError):
    pass


class DeclaracaoPdfNotFound(DeclaracaoPdfError):
    pass


def _storage_dir() -> Path:
    base_dir = os.environ.get("DECLARACAO_PDF_DIR")
    if base_dir:
        return Path(base_dir)
    return Path(tempfile.gettempdir()) / "secretariapadin_declaracoes_pdf"


def _safe_token(token: str) -> str:
    token = str(token or "").strip()
    if not PDF_TOKEN_RE.fullmatch(token):
        raise DeclaracaoPdfNotFound("PDF da declaracao nao encontrado.")
    return token


def _cleanup_old_files(directory: Path, now=None) -> None:
    now = now or time.time()
    try:
        for path in directory.glob("*.html"):
            try:
                if now - path.stat().st_mtime > PDF_TTL_SECONDS:
                    path.unlink(missing_ok=True)
            except OSError:
                continue
    except OSError:
        return


def prepare_html_for_pdf(html: str, app_root_path: str) -> str:
    static_uri = (Path(app_root_path) / "static").resolve().as_uri()
    return (
        html.replace('src="/static/', f'src="{static_uri}/')
        .replace("src='/static/", f"src='{static_uri}/")
        .replace('href="/static/', f'href="{static_uri}/')
        .replace("href='/static/", f"href='{static_uri}/")
    )


def create_declaracao_pdf_token(html: str, app_root_path: str) -> str:
    directory = _storage_dir()
    directory.mkdir(parents=True, exist_ok=True)
    _cleanup_old_files(directory)

    token = uuid4().hex
    html_path = directory / f"{token}.html"
    html_path.write_text(prepare_html_for_pdf(html, app_root_path), encoding="utf-8")
    return token


def _html_path_for_token(token: str) -> Path:
    token = _safe_token(token)
    html_path = _storage_dir() / f"{token}.html"
    if not html_path.exists():
        raise DeclaracaoPdfNotFound("PDF da declaracao expirou ou nao foi encontrado.")
    return html_path


def _candidate_browser_paths():
    env_paths = [
        os.environ.get("DECLARACAO_PDF_BROWSER_PATH"),
        os.environ.get("CHROME_PATH"),
        os.environ.get("CHROMIUM_PATH"),
    ]
    known_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
    ]
    return [path for path in [*env_paths, *known_paths] if path]


def find_browser_executable() -> str:
    for candidate in _candidate_browser_paths():
        if Path(candidate).exists():
            return candidate
    raise DeclaracaoPdfError(
        "Chrome/Chromium nao encontrado. Configure DECLARACAO_PDF_BROWSER_PATH "
        "com o caminho do navegador para gerar PDFs."
    )


def render_declaracao_pdf_bytes(token: str) -> bytes:
    html_path = _html_path_for_token(token)
    browser_path = find_browser_executable()

    with tempfile.TemporaryDirectory(prefix="secretariapadin_pdf_") as tmpdir:
        tmp_path = Path(tmpdir)
        pdf_path = tmp_path / "declaracao_escolar.pdf"
        user_data_dir = tmp_path / "chrome-profile"
        cmd = [
            browser_path,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            f"--user-data-dir={user_data_dir}",
            f"--print-to-pdf={pdf_path}",
            "--no-pdf-header-footer",
            html_path.resolve().as_uri(),
        ]

        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=60,
            check=False,
        )
        if result.returncode != 0 or not pdf_path.exists():
            details = (result.stderr or result.stdout or "").strip()
            raise DeclaracaoPdfError(
                "Nao foi possivel gerar o PDF da declaracao."
                + (f" Detalhes: {details}" if details else "")
            )

        return pdf_path.read_bytes()
