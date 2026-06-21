import os
import re
import uuid
from functools import wraps

from flask import (
    Blueprint,
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

from services.conferir_listas import (
    build_excel_report,
    format_turma_key,
    normalize_turma_lista,
    normalize_turma_sed,
    preview_sed_pdf_scope,
    load_result,
    prepare_result_for_view,
    run_conferencia,
    save_result,
)
from utils.uploads import allowed_excel_file, allowed_extension


confere_bp = Blueprint("confere", __name__, template_folder="templates")

UPLOAD_FOLDER = os.path.join(os.path.abspath(os.path.dirname(__file__)), "uploads")
RESULTS_FOLDER = os.path.join(UPLOAD_FOLDER, "conferencias")
ALLOWED_PDF_EXTENSIONS = {".pdf"}

os.makedirs(RESULTS_FOLDER, exist_ok=True)


def confere_login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login_route", next=request.url))
        return f(*args, **kwargs)

    return decorated_function


def _collect_sed_pdf_uploads():
    files = [file for file in request.files.getlist("sed_pdfs") if file and file.filename]
    items = []
    errors = []

    for file_storage in files:
        original_name = file_storage.filename or "arquivo_sem_nome.pdf"
        if not allowed_extension(original_name, ALLOWED_PDF_EXTENSIONS):
            errors.append(
                {
                    "arquivo": original_name,
                    "erro": "Arquivo rejeitado: envie apenas PDFs gerados pelo SED.",
                }
            )
            continue

        content = file_storage.read()
        if not content:
            errors.append(
                {
                    "arquivo": original_name,
                    "erro": "Arquivo PDF vazio ou nao enviado corretamente.",
                }
            )
            continue

        items.append({"filename": original_name, "content": content})

    return files, items, errors


def _attach_upload_errors(result, uploaded_files, upload_errors):
    if not upload_errors:
        return result

    result["pdf_errors"] = upload_errors + result.get("pdf_errors", [])
    summary = result.setdefault("summary", {})
    summary["total_pdfs_processados"] = len(uploaded_files)
    summary["total_pdfs_com_erro"] = len(result["pdf_errors"])
    return result


def _turma_sort_key(turma_key):
    match = re.fullmatch(r"(\d{1,2})([A-Z])", turma_key or "")
    if not match:
        return (999, turma_key or "")
    return (int(match.group(1)), match.group(2))


def _row_turma_key(row):
    return (
        normalize_turma_lista(row.get("turma_lista", ""))
        or normalize_turma_sed(row.get("turma_sed", ""))
        or "SEM_TURMA"
    )


def _build_result_groups(result):
    groups_by_key = {}
    category_order = {
        "inconsistencia_situacao": 0,
        "nao_encontrado_sed": 1,
        "nao_encontrado_lista": 2,
        "divergencia_cadastral": 3,
        "ok": 4,
    }

    for row in result.get("rows", []):
        turma_key = _row_turma_key(row)
        label = "Sem turma" if turma_key == "SEM_TURMA" else format_turma_key(turma_key)
        group = groups_by_key.setdefault(turma_key, {"key": turma_key, "label": label, "rows": []})
        group["rows"].append(row)

    groups = sorted(groups_by_key.values(), key=lambda group: _turma_sort_key(group["key"]))
    for group in groups:
        group["rows"].sort(
            key=lambda row: (
                category_order.get(row.get("categoria"), 99),
                row.get("nome_lista") or row.get("nome_sed") or "",
            )
        )
    return groups


@confere_bp.route("/", methods=["GET", "POST"])
@confere_login_required
def index():
    if request.method == "GET":
        return render_template("index.html", result=None)

    lista_piloto = request.files.get("lista_piloto")
    if not lista_piloto or not lista_piloto.filename:
        flash("Anexe a Lista Piloto em Excel para executar a conferencia.", "danger")
        return redirect(url_for("confere.index"))
    if not allowed_excel_file(lista_piloto.filename):
        flash("Formato invalido para Lista Piloto. Envie .xlsx, .xls ou .xlsm.", "danger")
        return redirect(url_for("confere.index"))

    uploaded_files, sed_items, upload_errors = _collect_sed_pdf_uploads()
    if not sed_items:
        if upload_errors:
            for error in upload_errors:
                flash(error["erro"], "warning")
        flash("Anexe pelo menos um PDF valido do SED para executar a conferencia.", "danger")
        return redirect(url_for("confere.index"))

    try:
        result = run_conferencia(lista_piloto, sed_items)
    except Exception as exc:
        flash(f"Nao foi possivel executar a conferencia: {exc}", "danger")
        return redirect(url_for("confere.index"))

    result = _attach_upload_errors(result, uploaded_files, upload_errors)
    result_id = uuid.uuid4().hex
    save_result(result, RESULTS_FOLDER, result_id)
    return redirect(url_for("confere.resultado", result_id=result_id))


@confere_bp.route("/preview-pdfs", methods=["POST"])
@confere_login_required
def preview_pdfs():
    uploaded_files, sed_items, upload_errors = _collect_sed_pdf_uploads()
    preview = (
        preview_sed_pdf_scope(sed_items)
        if sed_items
        else {
            "turmas": [],
            "turma_keys": [],
            "pdfs_validos": 0,
            "pdfs_ignorados": 0,
            "pdfs_selecionados": 0,
            "errors": [],
            "duplicate_files": [],
        }
    )
    errors = upload_errors + preview.get("errors", [])
    duplicate_files = preview.get("duplicate_files", [])
    return jsonify(
        {
            "turmas": preview.get("turmas", []),
            "pdfs_validos": preview.get("pdfs_validos", 0),
            "pdfs_ignorados": len(errors) + len(duplicate_files),
            "pdfs_selecionados": len(uploaded_files),
            "errors": errors,
            "duplicate_files": duplicate_files,
        }
    )


@confere_bp.route("/resultado/<result_id>", methods=["GET"])
@confere_login_required
def resultado(result_id):
    result = load_result(RESULTS_FOLDER, result_id)
    if result is None:
        flash("Resultado de conferencia nao encontrado. Execute a conferencia novamente.", "warning")
        return redirect(url_for("confere.index"))
    result = prepare_result_for_view(result)
    return render_template(
        "index.html",
        result=result,
        result_groups=_build_result_groups(result),
        result_id=result_id,
    )


@confere_bp.route("/exportar/excel/<result_id>", methods=["GET"])
@confere_login_required
def exportar_excel(result_id):
    result = load_result(RESULTS_FOLDER, result_id)
    if result is None:
        flash("Resultado de conferencia nao encontrado. Execute a conferencia novamente.", "warning")
        return redirect(url_for("confere.index"))

    output = build_excel_report(result)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="conferencia_lista_piloto_sed.xlsx",
    )


if __name__ == "__main__":
    app = Flask(__name__)
    app.secret_key = os.getenv("FLASK_SECRET_KEY") or uuid.uuid4().hex
    app.register_blueprint(confere_bp, url_prefix="/confere")
    app.run(debug=True)
