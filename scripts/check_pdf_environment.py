from pathlib import Path
import sys


APP_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP_ROOT))

from services.declaracoes_pdf import (
    create_declaracao_pdf_token,
    find_browser_executable,
    render_declaracao_pdf_bytes,
)


def main():
    browser_path = find_browser_executable()
    token = create_declaracao_pdf_token(
        """
        <!doctype html>
        <html lang="pt-br">
          <head>
            <meta charset="utf-8">
            <style>
              @page { size: A4; margin: 10mm 14mm; }
              body { font-family: Arial, sans-serif; font-size: 11pt; }
            </style>
          </head>
          <body>
            <h1>Teste de PDF</h1>
            <p>Ambiente pronto para gerar declarações em PDF.</p>
          </body>
        </html>
        """,
        str(APP_ROOT),
    )
    pdf_bytes = render_declaracao_pdf_bytes(token)
    print(f"OK: Chromium encontrado em {browser_path}")
    print(f"OK: PDF de teste gerado com {len(pdf_bytes)} bytes")


if __name__ == "__main__":
    main()
