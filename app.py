from flask import (
    Flask,
    request,
    redirect,
    url_for,
    render_template,
    jsonify,
    session,
    flash,
    current_app,
    send_file,
)
import os
import locale
from datetime import datetime
from functools import wraps
from io import BytesIO

import pandas as pd
import xlrd  # Para ler arquivos .xls, se necessário
from werkzeug.exceptions import RequestEntityTooLarge
from openpyxl import load_workbook, Workbook  # Usado para trabalhar com XLSX

from config import configure_app
from services.carteirinhas import (
    build_carteirinhas_context,
    mark_carteirinhas_impressas as _mark_carteirinhas_impressas,
)
from services.declaracoes import (
    DECLARACAO_PRINT_CSS,
    build_declaracao_escolar_context,
    build_lote_conclusao_5ano_context,
    build_lote_escolaridade_5ano_context,
    build_declaracao_personalizada_context,
    data_extenso_praia_grande,
    list_declaracao_alunos,
    load_declaracao_aluno_context,
)
from services.fotos import save_student_photo
from services.prazos import build_deadline_alerts as build_deadline_alerts_service
from services.quadros_atendimento import (
    ATENDIMENTO_CONFIG,
    AtendimentoMensalError,
    build_atendimento_mensal_file,
    normalize_mes_ref as _normalize_atendimento_mes_ref,
)
from services.quadros_inclusao import (
    QuantInclusaoError,
    build_quant_inclusao_file,
)
from services.quadros_quantitativo_mensal import (
    QuantitativoMensalError,
    build_quantitativo_mensal_file,
    get_default_mes_ano,
)
from services.quadros_transferencias import (
    TransferenciasError,
    build_transferencias_file,
)
from services.upload_sessions import save_excel_upload_to_session
from utils.dates import (
    parse_date_flexible,
    parse_period_date as _parse_period_date,
    parse_user_date as _parse_user_date,
)
from utils.excel import set_merged_cell_value
from utils.uploads import (
    save_excel_upload,
    validate_excel_upload,
)


# ==========================================================
#  CONFIGURAÇÃO BÁSICA DA APLICAÇÃO
# ==========================================================

# Tenta definir a localidade para formatação de datas em português
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except locale.Error:
    pass

app = Flask(__name__)
ACCESS_TOKEN = configure_app(app)


# Cria os diretórios necessários, se não existirem
os.makedirs("static/fotos", exist_ok=True)
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# Caminho do arquivo CSV relativo ao diretório do script
CSV_PATH = os.path.join(os.path.dirname(__file__), "uploads", "escolas.csv")

# Variável global para armazenar os dados do CSV
escolas_df = None


# ==========================================================
#  FUNÇÕES AUXILIARES – ESCOLAS (escolas.csv)
# ==========================================================

def carregar_escolas():
    """Carrega o CSV de escolas em um DataFrame global."""
    global escolas_df
    if os.path.exists(CSV_PATH):
        try:
            escolas_df = pd.read_csv(CSV_PATH, encoding="latin1", sep=";")
            print(f"[INFO] Arquivo {CSV_PATH} carregado com sucesso.")
        except Exception as e:
            escolas_df = None
            print(f"[ERRO] Falha ao carregar {CSV_PATH}: {e}")
    else:
        escolas_df = None
        print(f"[ERRO] Arquivo {CSV_PATH} não encontrado.")


def get_escolas_df():
    """Garante que o DataFrame de escolas está carregado."""
    global escolas_df
    if escolas_df is None or escolas_df.empty:
        print("[INFO] Recarregando arquivo escolas.csv...")
        carregar_escolas()
    return escolas_df


@app.before_request
def inicializar_escolas():
    """Garante que escolas.csv está carregado antes de cada requisição."""
    global escolas_df
    if escolas_df is None or (isinstance(escolas_df, pd.DataFrame) and escolas_df.empty):
        carregar_escolas()


@app.route("/escolas/search")
def escolas_search():
    """Endpoint para o Select2 buscar escolas no CSV."""
    df = get_escolas_df()
    query = request.args.get("q", "").lower().strip()
    results = []

    if df is not None and not df.empty and query:
        # Filtra usando pandas (assumindo coluna 3 = nome da escola)
        df_filtered = df[df.iloc[:, 3].str.lower().str.contains(query, na=False)]

        # Limita a 50 resultados para não sobrecarregar
        df_filtered = df_filtered.head(50)

        for _, row in df_filtered.iterrows():
            nome = str(row[3]).strip()
            municipio = str(row[2]).strip()
            uf = str(row[1]).strip()
            text = f"{nome} - {municipio}/{uf}"
            results.append({"id": nome, "text": text})

    return jsonify(results)


# Carrega o CSV na inicialização do sistema
carregar_escolas()


@app.errorhandler(RequestEntityTooLarge)
def handle_request_entity_too_large(_error):
    limit_mb = app.config.get("MAX_CONTENT_LENGTH", 0) // (1024 * 1024)
    message = f"Arquivo muito grande. O limite atual é de {limit_mb} MB."
    wants_json = request.accept_mimetypes.accept_json and not request.accept_mimetypes.accept_html
    if wants_json:
        return jsonify({"error": message}), 413
    flash(message, "error")
    return redirect(request.referrer or url_for("dashboard"))


# ==========================================================
#  BLUEPRINTS
# ==========================================================

from confere import confere_bp

app.register_blueprint(confere_bp, url_prefix="/confere")


# ==========================================================
#  HELPERS GERAIS
# ==========================================================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login_route", next=request.url))
        return f(*args, **kwargs)

    return decorated_function


def convert_xls_to_xlsx(file_like):
    """
    Converte um arquivo XLS (file-like) para um Workbook do openpyxl.
    """
    book_xlrd = xlrd.open_workbook(file_contents=file_like.read())
    wb = Workbook()
    # Remove a planilha padrão criada pelo openpyxl, se houver
    if "Sheet" in wb.sheetnames and len(book_xlrd.sheet_names()) > 0:
        std = wb.active
        wb.remove(std)

    for sheet_name in book_xlrd.sheet_names():
        sheet_xlrd = book_xlrd.sheet_by_name(sheet_name)
        ws = wb.create_sheet(title=sheet_name)
        for row in range(sheet_xlrd.nrows):
            for col in range(sheet_xlrd.ncols):
                ws.cell(row=row + 1, column=col + 1, value=sheet_xlrd.cell_value(row, col))

    return wb


def load_workbook_model(file):
    """
    Abre o arquivo do modelo XLSX (ou XLS convertendo-o para XLSX)
    preservando a formatação.
    """
    ext = os.path.splitext(file.filename)[1].lower()
    file.seek(0)
    if ext == ".xlsx":
        return load_workbook(file, data_only=False)
    elif ext == ".xls":
        content = file.read()
        return convert_xls_to_xlsx(BytesIO(content))
    else:
        raise ValueError("Formato de arquivo não suportado para o quadro modelo.")


def is_valid_plano(val):
    """Avalia se o valor do Plano de Ação é considerado válido."""
    if val is None:
        return False
    s = str(val).strip()
    return s not in ["", "-", "0", "#REF"]


# ==========================================================
#  (NECESSÁRIO) CARTEIRINHAS – MARCAR COMO IMPRESSAS
#  -> seu JS chama isso depois de imprimir (ou ao clicar em "Imprimir")
# ==========================================================

# Se você usa CSRFProtect (Flask-WTF), descomente a linha do csrf.exempt
# from flask_wtf.csrf import CSRFProtect
# csrf = CSRFProtect(app)

@app.route("/carteirinhas/marcar_impressas", methods=["POST"])
@login_required
# @csrf.exempt  # use isto se CSRF estiver bloqueando POST JSON
def marcar_carteirinhas_impressas():
    payload = request.get_json(silent=True) or {}
    return jsonify(_mark_carteirinhas_impressas(
        payload.get("rms", []),
        ano=payload.get("ano"),
    ))



# ==========================================================
#  DECLARAÇÕES – GERAÇÃO HTML (SINGULAR)
# ==========================================================

def gerar_declaracao_escolar(
    file_path,
    rm,
    tipo,
    file_path2=None,
    deve_historico=False,
    unidade_anterior=None,
    dados_frequencia=None,
):
    """
    Gera o HTML de uma declaração escolar (Escolaridade, Transferência, Conclusão
    ou Frequência) tanto para Fundamental quanto EJA, de acordo com
    session['declaracao_tipo'].

    file_path  -> caminho/arquivo padrão da lista piloto (salvo em sessão/ao entrar no sistema)
    file_path2 -> caminho/arquivo opcional, usado quando o usuário reenviar a lista
                  (por exemplo, após o servidor free acordar). SE informado, TERÁ PRIORIDADE.
    dados_frequencia -> dicionário opcional com os dados de frequência por mês,
                        utilizado apenas quando tipo == "Frequencia".
    """
    global escolas_df

    # Se um segundo caminho foi informado (lista reenviada), ele tem prioridade.
    effective_path = file_path2 if file_path2 is not None else file_path
    if file_path2 is not None:
        print("[DEBUG] gerar_declaracao_escolar: usando file_path2 =", effective_path)
    else:
        print("[DEBUG] gerar_declaracao_escolar: usando file_path  =", effective_path)

    segmento_declaracao = session.get("declaracao_tipo")
    aluno_context = load_declaracao_aluno_context(
        effective_path,
        rm,
        segmento_declaracao,
        tipo,
    )
    if aluno_context is None:
        return None

    context = build_declaracao_escolar_context(
        tipo=tipo,
        segmento=segmento_declaracao,
        nome=aluno_context["nome"],
        ra=aluno_context["ra"],
        ra_label=aluno_context["ra_label"],
        data_nasc=aluno_context["data_nasc"],
        serie=aluno_context["serie"],
        horario=aluno_context["horario"],
        semestre_texto=aluno_context["semestre_texto"],
        row=aluno_context["row"],
        notas_tabela_html=aluno_context["notas_tabela_html"],
        deve_historico=deve_historico,
        unidade_anterior=unidade_anterior,
        escolas_df=escolas_df,
        dados_frequencia=dados_frequencia,
    )
    if context is None:
        return None

    data_extenso_str = data_extenso_praia_grande()
    additional_css = DECLARACAO_PRINT_CSS

    return render_template(
        "declaracao_print.html",
        titulo=context["titulo"],
        data_extenso=data_extenso_str,
        declaracao_text=context["declaracao_text"],
        additional_css=additional_css,
        body_classes=context["body_classes"],
        print_body_padding="0.5cm 0.5cm",
    )


# ==========================================================
#  DECLARAÇÃO PERSONALIZADA (Fundamental / EJA)
# ==========================================================

def gerar_declaracao_personalizada(dados):
    """
    Gera o HTML de declarações personalizadas (Conclusão, Matrícula cancelada
    ou Não Comparecimento - NCOM).
    """
    context = build_declaracao_personalizada_context(dados)
    if context is None:
        return None

    data_extenso_str = data_extenso_praia_grande()
    additional_css = DECLARACAO_PRINT_CSS

    return render_template(
        "declaracao_print.html",
        titulo=context["titulo"],
        data_extenso=data_extenso_str,
        declaracao_text=context["declaracao_text"],
        additional_css=additional_css,
        body_classes=[],
        print_body_padding="1.5cm 1.5cm",
    )


# ==========================================================
#  AUTENTICAÇÃO / DASHBOARD
# ==========================================================

@app.route("/login", methods=["GET", "POST"])
def login_route():
    error = None
    if request.method == "POST":
        if not ACCESS_TOKEN:
            error = "ACCESS_TOKEN não configurado no servidor. Configure a variável de ambiente antes de usar o sistema."
            return render_template("login.html", error=error), 503

        token = request.form.get("token")
        if token == ACCESS_TOKEN:
            session["logged_in"] = True

            # ==========================================================
            # AJUSTE: EJA NÃO É MAIS OBRIGATÓRIA
            # - Só exige lista_fundamental para seguir para o dashboard
            # - lista_eja permanece suportada, mas opcional
            # ==========================================================
            if "lista_fundamental" not in session:
                return redirect(url_for("upload_listas"))

            return redirect(url_for("dashboard"))
        else:
            error = "Token inválido. Tente novamente."

    return render_template("login.html", error=error)


@app.route("/logout")
def logout_route():
    session.clear()
    return redirect(url_for("login_route"))


@app.route("/upload_listas", methods=["GET", "POST"])
@login_required
def upload_listas():
    if request.method == "POST":
        fundamental_file = request.files.get("lista_fundamental")
        eja_file = request.files.get("lista_eja")  # agora é opcional

        if not fundamental_file or fundamental_file.filename == "":
            flash("Selecione a Lista Piloto - REGULAR - 2025", "error")
            return redirect(url_for("upload_listas"))

        valid, message = validate_excel_upload(fundamental_file)
        if not valid:
            flash(message, "error")
            return redirect(url_for("upload_listas"))

        valid, message = validate_excel_upload(eja_file, required=False)
        if not valid:
            flash(f"Lista EJA inválida: {message}", "error")
            return redirect(url_for("upload_listas"))

        # Salva Fundamental (obrigatório)
        save_excel_upload_to_session(
            fundamental_file,
            session,
            "lista_fundamental",
            app.config["UPLOAD_FOLDER"],
            prefix="fundamental",
        )

        # Salva EJA (opcional)
        eja_salva = False
        if eja_file and eja_file.filename:
            save_excel_upload_to_session(
                eja_file,
                session,
                "lista_eja",
                app.config["UPLOAD_FOLDER"],
                prefix="eja",
            )
            eja_salva = True

        if eja_salva:
            flash("Listas carregadas com sucesso (Fundamental e EJA).", "success")
        else:
            flash(
                "Lista do Fundamental carregada com sucesso. A lista de EJA é opcional e não foi enviada.",
                "success",
            )

        return redirect(url_for("dashboard"))

    return render_template("upload_listas.html")


@app.route("/", methods=["GET"])
@login_required
def dashboard():
    return render_template("dashboard.html")


# ==========================================================
#  CARTEIRINHAS – ROTA PRINCIPAL (AJUSTADA)
#  - Mantém tudo que já existe
#  - Adiciona suporte ao filtro "somente_nao_impressas"
#  - Persiste ambos filtros em session p/ GET
# ==========================================================

@app.route("/carteirinhas", methods=["GET", "POST"])
@login_required
def carteirinhas():
    if request.method == "POST":
        file_path = None

        # Checkbox no form (ex.: <input type="checkbox" name="somente_com_foto">)
        somente_com_foto = request.form.get("somente_com_foto") in ("1", "on", "true", "True", "SIM", "sim")
        session["carteirinhas_somente_com_foto"] = somente_com_foto

        # NOVO: checkbox "somente_nao_impressas"
        # (ex.: <input type="checkbox" name="somente_nao_impressas">)
        somente_nao_impressas = request.form.get("somente_nao_impressas") in ("1", "on", "true", "True", "SIM", "sim")
        session["carteirinhas_somente_nao_impressas"] = somente_nao_impressas

        if "excel_file" in request.files and request.files["excel_file"].filename != "":
            file = request.files["excel_file"]
            valid, message = validate_excel_upload(file)
            if not valid:
                flash(message, "error")
                return redirect(url_for("carteirinhas"))

            file_path = save_excel_upload_to_session(
                file,
                session,
                "lista_fundamental",
                app.config["UPLOAD_FOLDER"],
                prefix="carteirinhas",
            )
        else:
            file_path = session.get("lista_fundamental")

        if not file_path or not os.path.exists(file_path):
            flash("Nenhum arquivo selecionado. Envie a lista piloto do Fundamental.", "info")
            return redirect(url_for("carteirinhas"))

        flash("Gerando carteirinhas. Aguarde...", "info")
        context = build_carteirinhas_context(
            file_path,
            somente_com_foto=somente_com_foto,
            somente_nao_impressas=somente_nao_impressas,  # NOVO
            ano=datetime.now().year,                      # NOVO (para log por ano)
        )
        return render_template("gerar_carteirinhas.html", **context)

    # GET: passa o estado atual dos filtros para o template marcar os checkbox
    somente_com_foto = session.get("carteirinhas_somente_com_foto", False)
    somente_nao_impressas = session.get("carteirinhas_somente_nao_impressas", False)

    return render_template(
        "carteirinhas.html",
        somente_com_foto=somente_com_foto,
        somente_nao_impressas=somente_nao_impressas,  # NOVO
    )


#  DECLARAÇÕES – CONCLUSÃO 5º ANO (LOTE)
# ==========================================================

@app.route("/declaracao/conclusao_5ano")
@login_required
def declaracao_conclusao_5ano():
    if session.get("declaracao_tipo") != "Fundamental":
        flash(
            "As declarações em lote de 5º ano estão disponíveis apenas para o Fundamental.",
            "error",
        )
        return redirect(url_for("declaracao_tipo"))

    file_path = session.get("declaracao_excel") or session.get("lista_fundamental")

    if not file_path or not os.path.exists(file_path):
        flash(
            "Arquivo Excel do Fundamental não encontrado. "
            "Anexe a lista piloto novamente pela tela de declarações.",
            "error",
        )
        return redirect(url_for("declaracao_tipo", segmento="Fundamental"))

    data_extenso_str = current_app.config.get(
        "CONCLUSAO_5ANO_DATE_TEXT",
        f"Praia Grande, 22 de dezembro de {datetime.now().year}",
    )
    context = build_lote_conclusao_5ano_context(
        file_path,
        data_extenso_str,
        bolsa_familia_src=url_for("static", filename="logos/bolsa_familia.jpg"),
    )
    if not context["registros"]:
        flash("Nenhum aluno de 5º ano encontrado na lista piloto.", "error")
        return redirect(url_for("declaracao_tipo", segmento="Fundamental"))

    return render_template(
        "declaracao_conclusao_5ano.html",
        registros=context["registros"],
        data_extenso=context["data_extenso"],
        titulo=context["titulo"],
        total=context["total"],
    )


# ==========================================================
#  DECLARAÇÕES – TELA ÚNICA (Fundamental / EJA / Personalizada)
# ==========================================================

@app.route("/declaracao/tipo", methods=["GET", "POST"])
@login_required
def declaracao_tipo():
    if request.method == "POST":
        modo_declaracao = request.form.get("modo_declaracao")

        # -------------------------------
        # FLUXO: DECLARAÇÃO PERSONALIZADA
        # -------------------------------
        if modo_declaracao == "personalizada":
            segmento_pers = request.form.get("segmento_personalizado")
            if segmento_pers not in ("Fundamental", "EJA"):
                flash(
                    "Selecione o segmento (Ensino Fundamental ou EJA) na declaração personalizada.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

            nome_aluno = (request.form.get("nome_aluno") or "").strip()
            data_nascimento = request.form.get("data_nascimento")
            ra = (request.form.get("ra") or "").strip()
            tipo_pers = request.form.get("tipo_declaracao_personalizada")

            if (
                not nome_aluno
                or not data_nascimento
                or not ra
                or tipo_pers not in ("Conclusao", "MatriculaCancelada", "NCOM")
            ):
                flash(
                    "Preencha todos os dados obrigatórios da declaração personalizada.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

            dados_personalizados = {
                "segmento": segmento_pers,
                "nome_aluno": nome_aluno,
                "data_nascimento": data_nascimento,
                "ra": ra,
                "tipo_declaracao": tipo_pers,
            }

            if tipo_pers == "Conclusao":
                ano_serie_concluida = (request.form.get("ano_serie_concluida") or "").strip()
                ano_conclusao = (request.form.get("ano_conclusao") or "").strip()
                deve_hist_unidade = request.form.get("deve_historico_unidade")
                semestre_conclusao = (request.form.get("semestre_conclusao") or "").strip()

                campos_invalidos = (
                    not ano_serie_concluida
                    or not ano_conclusao
                    or deve_hist_unidade not in ("Sim", "Não")
                )

                if segmento_pers == "EJA" and not semestre_conclusao:
                    campos_invalidos = True

                if campos_invalidos:
                    flash(
                        "Preencha todos os campos da declaração personalizada de conclusão "
                        "(para EJA é obrigatório informar o semestre).",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

                dados_personalizados.update(
                    {
                        "ano_serie_concluida": ano_serie_concluida,
                        "ano_conclusao": ano_conclusao,
                        "deve_historico_unidade": (deve_hist_unidade == "Sim"),
                        "semestre_conclusao": semestre_conclusao,
                    }
                )

            elif tipo_pers == "MatriculaCancelada":
                ano_serie_matricula = (request.form.get("ano_serie_matricula") or "").strip()
                ano_matricula = (request.form.get("ano_matricula") or "").strip()
                semestre_matricula = (request.form.get("semestre_matricula") or "").strip()

                if not ano_serie_matricula or not ano_matricula or not semestre_matricula:
                    flash(
                        "Preencha todos os campos da declaração de matrícula cancelada.",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

                dados_personalizados.update(
                    {
                        "ano_serie_matricula": ano_serie_matricula,
                        "ano_matricula": ano_matricula,
                        "semestre_matricula": semestre_matricula,
                    }
                )

            elif tipo_pers == "NCOM":
                ano_serie_vaga = (request.form.get("ano_serie_vaga") or "").strip()
                ano_referencia_ncom = (request.form.get("ano_referencia_ncom") or "").strip()
                semestre_referencia_ncom = (
                    request.form.get("semestre_referencia_ncom") or ""
                ).strip()

                if not ano_serie_vaga or not ano_referencia_ncom:
                    flash(
                        "Preencha todos os campos obrigatórios da declaração de Não Comparecimento (NCOM).",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

                dados_personalizados.update(
                    {
                        "ano_serie_vaga": ano_serie_vaga,
                        "ano_referencia_ncom": ano_referencia_ncom,
                        "semestre_referencia_ncom": semestre_referencia_ncom,
                    }
                )

            declaracao_html = gerar_declaracao_personalizada(dados_personalizados)

            if declaracao_html is None:
                flash(
                    "Não foi possível gerar a declaração personalizada. Verifique os dados informados.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

            return declaracao_html

        # -------------------------------
        # FLUXO NORMAL: FUNDAMENTAL / EJA
        # -------------------------------
        segmento = request.form.get("segmento_escolhido")
        if segmento not in ("Fundamental", "EJA"):
            flash("Selecione se a declaração é do Fundamental ou EJA antes de gerar.", "error")
            return redirect(url_for("declaracao_tipo"))

        rm = (request.form.get("rm") or "").strip()
        tipo = (request.form.get("tipo") or "").strip()

        tipo_lower = tipo.lower()
        if tipo_lower in ("transferencia", "transferência"):
            tipo = "Transferencia"
        elif tipo_lower in ("conclusao", "conclusão"):
            tipo = "Conclusão"
        elif tipo_lower in ("frequencia", "frequência"):
            tipo = "Frequencia"

        deve_historico_str = request.form.get("deve_historico")

        unidade_select = (request.form.get("unidade_anterior_select") or "").strip()
        unidade_manual = (request.form.get("unidade_anterior_manual") or "").strip()
        unidade_anterior = unidade_select or unidade_manual

        file_path = None
        excel_file = request.files.get("excel_file")
        novo_upload = excel_file is not None and excel_file.filename

        if novo_upload:
            valid, message = validate_excel_upload(excel_file)
            if not valid:
                flash(message, "error")
                return redirect(url_for("declaracao_tipo", segmento=segmento))

            session_key = "lista_fundamental" if segmento == "Fundamental" else "lista_eja"
            file_path = save_excel_upload_to_session(
                excel_file,
                session,
                session_key,
                app.config["UPLOAD_FOLDER"],
                prefix="declaracao",
            )
        else:
            if segmento == "Fundamental":
                file_path = session.get("lista_fundamental")
            else:
                file_path = session.get("lista_eja")

        if not file_path or not os.path.exists(file_path):
            flash(
                "Nenhuma lista piloto encontrada para este segmento. Anexe o arquivo em Excel.",
                "error",
            )
            return redirect(url_for("declaracao_tipo", segmento=segmento))

        session["declaracao_tipo"] = segmento
        session["declaracao_excel"] = file_path

        if novo_upload and (not rm or not tipo):
            flash(
                "Lista piloto carregada com sucesso. Agora selecione o aluno e o tipo de declaração.",
                "success",
            )
            return redirect(url_for("declaracao_tipo", segmento=segmento))

        if not rm or not tipo:
            flash("Escolha o aluno e o tipo de declaração.", "error")
            return redirect(url_for("declaracao_tipo", segmento=segmento))

        if tipo in ("Transferencia", "Conclusão"):
            if deve_historico_str not in ("sim", "nao"):
                flash("Por favor, responda se o aluno deve o histórico escolar.", "error")
                return redirect(url_for("declaracao_tipo", segmento=segmento))

            if deve_historico_str == "sim" and not unidade_anterior:
                flash(
                    "Informe a unidade escolar anterior para a qual o aluno deve o histórico.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento=segmento))

            deve_historico = deve_historico_str == "sim"
        else:
            deve_historico = False
            unidade_anterior = ""

        dados_frequencia = None

        if tipo == "Frequencia":
            meses = [
                ("jan", "Janeiro"),
                ("fev", "Fevereiro"),
                ("mar", "Março"),
                ("abr", "Abril"),
                ("mai", "Maio"),
                ("jun", "Junho"),
                ("jul", "Julho"),
                ("ago", "Agosto"),
                ("set", "Setembro"),
                ("out", "Outubro"),
                ("nov", "Novembro"),
                ("dez", "Dezembro"),
            ]

            dados_frequencia = {"meses": []}
            algum_valido = False

            for mes_id, mes_nome in meses:
                dias_raw = (request.form.get(f"freq_{mes_id}_dias") or "").strip()
                faltas_raw = (request.form.get(f"freq_{mes_id}_faltas") or "").strip()

                if not dias_raw and not faltas_raw:
                    dados_frequencia["meses"].append(
                        {
                            "id": mes_id,
                            "nome": mes_nome,
                            "dias_letivos": None,
                            "faltas": None,
                            "frequencia": None,
                            "preenchido": False,
                        }
                    )
                    continue

                try:
                    dias = float(dias_raw.replace(",", ".")) if dias_raw else None
                    faltas = float(faltas_raw.replace(",", ".")) if faltas_raw else None
                except ValueError:
                    flash(
                        "Verifique os valores de dias letivos e faltas informados na frequência.",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento=segmento))

                if dias is None or faltas is None:
                    flash(
                        "Para cada mês de frequência preenchido, informe tanto os dias letivos quanto as faltas.",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento=segmento))

                if dias <= 0 or faltas < 0 or faltas > dias:
                    flash(
                        "Os valores de dias letivos e faltas são inválidos em um ou mais meses. "
                        "Verifique e tente novamente.",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento=segmento))

                freq_percent = ((dias - faltas) / dias) * 100.0
                algum_valido = True

                dados_frequencia["meses"].append(
                    {
                        "id": mes_id,
                        "nome": mes_nome,
                        "dias_letivos": dias,
                        "faltas": faltas,
                        "frequencia": round(freq_percent, 1),
                        "preenchido": True,
                    }
                )

            if not algum_valido:
                flash(
                    "Informe ao menos um mês de frequência com dias letivos e faltas válidos.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento=segmento))

        declaracao_html = gerar_declaracao_escolar(
            file_path=file_path,
            rm=rm,
            tipo=tipo,
            deve_historico=deve_historico,
            unidade_anterior=unidade_anterior,
            dados_frequencia=dados_frequencia,
        )

        if declaracao_html is None:
            flash("Aluno não encontrado na lista piloto.", "error")
            return redirect(url_for("declaracao_tipo", segmento=segmento))

        return declaracao_html

    # --------------------------------------
    # GET: EXIBE A TELA
    # --------------------------------------
    segmento = request.args.get("segmento")
    if segmento not in ("Fundamental", "EJA", "Personalizado"):
        segmento = None

    alunos = []
    tem_lista = False

    if segmento == "Fundamental":
        file_path = session.get("lista_fundamental")
        if file_path and os.path.exists(file_path):
            tem_lista = True
            session["declaracao_tipo"] = "Fundamental"
            session["declaracao_excel"] = file_path

            alunos = list_declaracao_alunos(file_path, "Fundamental")

    elif segmento == "EJA":
        # EJA permanece suportada, porém agora é opcional no sistema como um todo.
        file_path = session.get("lista_eja")
        if file_path and os.path.exists(file_path):
            tem_lista = True
            session["declaracao_tipo"] = "EJA"
            session["declaracao_excel"] = file_path

            alunos = list_declaracao_alunos(file_path, "EJA")
        else:
            tem_lista = False
            alunos = []

    dashboard_url = url_for("dashboard")
    conclusao_5ano_url = url_for("declaracao_conclusao_5ano")
    escolaridade_5ano_url = url_for("declaracao_escolaridade_5ano")

    return render_template(
        "declaracao_tipo.html",
        segmento=segmento,
        tem_lista=tem_lista,
        alunos=alunos,
        dashboard_url=dashboard_url,
        conclusao_5ano_url=conclusao_5ano_url,
        escolaridade_5ano_url=escolaridade_5ano_url,
    )


@app.route("/declaracao/escolaridade_5ano")
@login_required
def declaracao_escolaridade_5ano(file_path_arg=None):
    if file_path_arg:
        file_path = file_path_arg
    else:
        file_path = session.get("declaracao_excel")

    if session.get("declaracao_tipo") != "Fundamental":
        flash(
            "As declarações de escolaridade de 5º ano só podem ser geradas "
            "com a lista piloto do Ensino Fundamental.",
            "error",
        )
        return redirect(url_for("declaracao_tipo", segmento="Fundamental"))

    if not file_path or not os.path.exists(file_path):
        flash(
            "Nenhuma lista piloto do Ensino Fundamental está carregada. "
            "Anexe a lista piloto novamente para gerar as declarações.",
            "error",
        )
        return redirect(url_for("declaracao_tipo", segmento="Fundamental"))

    context = build_lote_escolaridade_5ano_context(file_path)

    if not context["registros"]:
        flash(
            "Nenhum aluno de 5º ano foi encontrado na lista piloto para "
            "gerar as declarações de escolaridade.",
            "error",
        )
        return redirect(url_for("declaracao_tipo", segmento="Fundamental"))

    return render_template(
        "declaracao_escolaridade_5ano.html",
        registros=context["registros"],
        data_extenso=context["data_extenso"],
        titulo=context["titulo"],
    )


# ==========================================================
#  UPLOAD DE FOTOS (CARTEIRINHAS)
# ==========================================================

@app.route("/upload_foto", methods=["POST"])
@login_required
def upload_foto():
    rm = (request.form.get("rm") or "").strip()
    if not rm:
        flash("RM não fornecido.", "error")
        return redirect(url_for("carteirinhas"))

    file = request.files.get("foto_file")
    try:
        save_student_photo(file, rm)
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("carteirinhas"))

    flash("Foto anexada com sucesso.", "success")
    return redirect(url_for("carteirinhas"))


@app.route("/upload_multiplas_fotos", methods=["POST"])
@login_required
def upload_multiplas_fotos():
    rms = request.form.getlist("rm[]")
    files = request.files.getlist("foto_file[]")

    if not files:
        flash("Nenhuma foto enviada.", "error")
        return redirect(url_for("carteirinhas"))

    total_salvas = 0
    total_ignoradas = 0

    for rm, file in zip(rms, files):
        rm = (rm or "").strip()
        try:
            save_student_photo(file, rm)
        except ValueError:
            total_ignoradas += 1
            continue

        total_salvas += 1

    if total_salvas:
        msg = f"Foto(s) anexada(s) com sucesso: {total_salvas} arquivo(s)."
        if total_ignoradas:
            msg += f" {total_ignoradas} arquivo(s) foram ignorados por RM ou formato inválido."
        flash(msg, "success")
    else:
        flash("Nenhuma foto válida foi enviada.", "error")

    return redirect(url_for("carteirinhas"))


@app.route("/upload_inline_foto", methods=["POST"])
@login_required
def upload_inline_foto():
    file = request.files.get("foto_file")
    rm = (request.form.get("rm") or "").strip()

    if not rm:
        return jsonify({"error": "RM não fornecido"}), 400

    try:
        saved = save_student_photo(file, rm)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    return jsonify({"url": saved["url"], "message": "Foto anexada com sucesso"}), 200


# ==========================================================
#  QUADROS – HELPERS COMUNS + EJA (LIGA/DESLIGA)
#  (FIX RÁPIDO E SEGURO: sem colisão de nomes, sem sobrescrita de helpers)
# ==========================================================

# ----------------------------------------------------------
# EJA: liga/desliga (FIX)
# ----------------------------------------------------------
def _is_eja_enabled() -> bool:
    """
    Desativada por padrão:
      - Ative via variável de ambiente: ENABLE_EJA=1
      - Ou via config: app.config["ENABLE_EJA"] = True

    FIX: remove duplicidade/override incorreto.
    """
    def _to_bool(v) -> bool:
        if v is None:
            return False
        if isinstance(v, bool):
            return v
        s = str(v).strip().lower()
        return s in ("1", "true", "t", "yes", "y", "on")

    # 1) prioridade para config do Flask
    cfg = None
    try:
        cfg = current_app.config.get("ENABLE_EJA", None)
    except Exception:
        cfg = None

    if cfg is not None:
        return _to_bool(cfg)

    # 2) fallback: env var
    return _to_bool(os.getenv("ENABLE_EJA", "0"))


def _save_upload_to_session(file_storage, session_key: str, prefix: str) -> str:
    """
    Salva arquivo enviado em UPLOAD_FOLDER e grava em session[session_key].
    Retorna o path salvo.
    """
    upload_folder = None
    try:
        upload_folder = current_app.config.get("UPLOAD_FOLDER")
    except Exception:
        upload_folder = None

    if not upload_folder:
        # mantém compatibilidade com seu app.py original (que define app.config["UPLOAD_FOLDER"])
        upload_folder = "uploads"

    return save_excel_upload_to_session(
        file_storage,
        session,
        session_key,
        upload_folder,
        prefix=prefix,
    )


# ==========================================================
#  QUADROS – MENU PRINCIPAL
# ==========================================================
@app.route("/quadros")
@login_required
def quadros():
    return render_template("quadros.html")


# ==========================================================
#  QUADRO – INCLUSÃO (DESATIVADO / ARQUIVADO)
# ==========================================================
@app.route("/quadros/inclusao", methods=["GET", "POST"])
@login_required
def quadros_inclusao():
    flash("Quadro de Inclusão foi migrado para Drive e está desativado no sistema no momento.", "info")
    return redirect(url_for("quadros"))


# ==========================================================
#  QUADRO – ATENDIMENTO MENSAL (CORRIGIDO / FUTURE-PROOF)
# ==========================================================

@app.route('/quadros/atendimento_mensal', methods=['GET', 'POST'])
@login_required
def quadro_atendimento_mensal():
    if request.method == 'POST':
        enable_eja = _is_eja_enabled()

        responsavel = (request.form.get("responsavel") or "").strip()
        rf = (request.form.get("rf") or "").strip()

        mes_ref = _normalize_atendimento_mes_ref(request.form.get("mes_ref"))

        fundamental_file = request.files.get('lista_fundamental')
        eja_file = request.files.get('lista_eja')

        try:
            if fundamental_file and fundamental_file.filename != '':
                _save_upload_to_session(fundamental_file, 'lista_fundamental', prefix='atendimento')

            if enable_eja and eja_file and eja_file.filename != '':
                _save_upload_to_session(eja_file, 'lista_eja', prefix='atendimento_eja')
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        file_path = session.get('lista_fundamental')
        if not file_path or not os.path.exists(file_path):
            flash("Nenhum arquivo da Lista Piloto FUNDAMENTAL disponível.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        model_path = os.path.join("modelos", "Quadro de Atendimento Mensal - Modelo.xlsx")
        eja_path = session.get('lista_eja') if enable_eja else None

        try:
            result = build_atendimento_mensal_file(
                fundamental_path=file_path,
                model_path=model_path,
                responsavel=responsavel,
                rf=rf,
                mes_ref=mes_ref,
                enable_eja=enable_eja,
                eja_path=eja_path,
            )
        except AtendimentoMensalError as e:
            flash(str(e), "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        debug_log = result.debug_log
        for warning in result.warnings:
            flash(warning, "warning")

        if ATENDIMENTO_CONFIG["ENABLE_DEBUG_LOG"]:
            try:
                current_app.logger.info("=== DEBUG ATENDIMENTO MENSAL ===\n%s", "\n".join(debug_log))
            except Exception:
                print("=== DEBUG ATENDIMENTO MENSAL ===")
                print("\n".join(debug_log))

        return send_file(
            result.output,
            as_attachment=True,
            download_name=result.filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return render_template('quadro_atendimento_mensal.html', enable_eja=_is_eja_enabled())


# ==========================================================
#  QUADRO – TRANSFERÊNCIAS (FIX H/I/J + ALERTAS DE FALTA DE INFORMAÇÃO)
# ==========================================================

@app.route("/quadros/transferencias", methods=["GET", "POST"])
@login_required
def quadro_transferencias():
    if request.method == "POST":
        enable_eja = _is_eja_enabled()

        period_start_str = request.form.get("period_start")
        period_end_str = request.form.get("period_end")

        responsavel = request.form.get("responsavel")
        diretor_nome = request.form.get("diretor_nome") or "Luciana Rocha Augustinho"
        data_quadro_in = request.form.get("data_quadro")

        fundamental_file = request.files.get("lista_fundamental")
        eja_file = request.files.get("lista_eja")

        if not period_start_str or not period_end_str or not responsavel:
            flash("Por favor, preencha todos os campos.", "error")
            return redirect(url_for("quadro_transferencias"))

        if fundamental_file and fundamental_file.filename != "":
            valid, message = validate_excel_upload(fundamental_file)
            if not valid:
                flash(message, "error")
                return redirect(url_for("quadro_transferencias"))

            upload_folder = current_app.config.get("UPLOAD_FOLDER", "uploads")
            fundamental_path = save_excel_upload_to_session(
                fundamental_file,
                session,
                "lista_fundamental",
                upload_folder,
                prefix="fundamental",
            )
        else:
            fundamental_path = session.get("lista_fundamental")
            if not fundamental_path or not os.path.exists(fundamental_path):
                flash("Lista Piloto Fundamental não encontrada.", "error")
                return redirect(url_for("quadro_transferencias"))

        eja_path = None
        if enable_eja and eja_file and eja_file.filename != "":
            valid, message = validate_excel_upload(eja_file)
            if not valid:
                flash(f"Lista EJA inválida: {message}", "error")
                return redirect(url_for("quadro_transferencias"))

            upload_folder = current_app.config.get("UPLOAD_FOLDER", "uploads")
            eja_path = save_excel_upload_to_session(
                eja_file,
                session,
                "lista_eja",
                upload_folder,
                prefix="eja",
            )
        elif enable_eja:
            eja_path = session.get("lista_eja")

        try:
            period_start = _parse_period_date(period_start_str, "a data inicial")
            period_end = _parse_period_date(period_end_str, "a data final")
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("quadro_transferencias"))

        if period_end < period_start:
            flash("A data final não pode ser menor que a data inicial.", "error")
            return redirect(url_for("quadro_transferencias"))

        model_path = os.path.join("modelos", "Quadro Informativo - Modelo.xlsx")
        data_quadro_dt = _parse_user_date(data_quadro_in) or datetime.now()

        try:
            result = build_transferencias_file(
                fundamental_path=fundamental_path,
                model_path=model_path,
                period_start=period_start,
                period_end=period_end,
                responsavel=responsavel,
                diretor_nome=diretor_nome,
                data_quadro_dt=data_quadro_dt,
                enable_eja=enable_eja,
                eja_path=eja_path,
            )
        except TransferenciasError as e:
            flash(str(e), "error")
            return redirect(url_for("quadro_transferencias"))

        try:
            current_app.logger.info("\n".join(result.debug))
        except Exception:
            print("\n".join(result.debug))

        resp = send_file(
            result.output,
            as_attachment=True,
            download_name=result.filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Dados pessoais ficam na aba ALERTAS do Excel, não em headers HTTP.
        if result.missing_info_alerts:
            resp.headers["X-Transferencias-MissingInfo-Count"] = str(len(result.missing_info_alerts))

        return resp

    return render_template("quadro_transferencias.html")


# ==========================================================
#  QUADRO – QUANTITATIVO MENSAL (Fundamental)
#  (mantido: parse flexível do período + debug sheet oculta)
# ==========================================================

@app.route("/quadros/quantitativo_mensal", methods=["GET", "POST"])
@login_required
def quadro_quantitativo_mensal():
    if request.method == "POST":
        period_start_str = request.form.get("period_start")
        period_end_str = request.form.get("period_end")
        responsavel = request.form.get("responsavel")
        mes_ano = request.form.get("mes_ano")

        if not responsavel or not str(responsavel).strip():
            flash("Preencha o campo Responsável.", "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        default_year = datetime.now().year

        try:
            period_start = parse_date_flexible(period_start_str, default_year=default_year, field_label="a data inicial")
            period_end = parse_date_flexible(period_end_str, default_year=default_year, field_label="a data final")
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        if period_end < period_start:
            flash("A data final não pode ser menor que a data inicial.", "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        if not mes_ano or not str(mes_ano).strip():
            mes_ano = get_default_mes_ano()
        mes_ano = str(mes_ano).strip()

        fundamental_file = request.files.get("lista_fundamental")
        if fundamental_file and fundamental_file.filename:
            valid, message = validate_excel_upload(fundamental_file)
            if not valid:
                flash(message, "error")
                return redirect(url_for("quadro_quantitativo_mensal"))

            upload_folder = current_app.config.get("UPLOAD_FOLDER", "uploads")
            fundamental_path = save_excel_upload_to_session(
                fundamental_file,
                session,
                "lista_fundamental",
                upload_folder,
                prefix="fundamental",
            )
        else:
            fundamental_path = session.get("lista_fundamental")
            if not fundamental_path or not os.path.exists(fundamental_path):
                flash("Arquivo da Lista Piloto Fundamental não encontrado.", "error")
                return redirect(url_for("quadro_quantitativo_mensal"))

        model_path = os.path.join("modelos", "Quadro Quantitativo Mensal - Modelo.xlsx")
        ano_letivo = int(current_app.config.get("SCHOOL_YEAR", datetime.now().year))

        try:
            result = build_quantitativo_mensal_file(
                fundamental_path=fundamental_path,
                model_path=model_path,
                period_start=period_start,
                period_end=period_end,
                responsavel=responsavel,
                mes_ano=mes_ano,
                default_year=default_year,
                ano_letivo=ano_letivo,
            )
        except QuantitativoMensalError as e:
            flash(str(e), "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        current_app.logger.info(
            "[QUADRO_QUANTITATIVO] periodo=%s..%s | counted=%s | discarded=%s | ano_padrao_sem_ano=%s",
            period_start.strftime("%d/%m/%Y"),
            period_end.strftime("%d/%m/%Y"),
            result.counted,
            result.discarded,
            result.default_year,
        )

        return send_file(
            result.output,
            as_attachment=True,
            download_name=result.filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    return render_template("quadro_quantitativo_mensal.html", default_mes_ano=get_default_mes_ano())

# ==========================================================
#  QUADRO QUANTITATIVO DE INCLUSÃO – REGULAR (EJA DESCONSIDERADA)
#  - Contagem por turma via LISTA CORRIDA
#  - Regras:
#      * Turma: Coluna A
#      * RM (identificador único): Coluna C (NÃO contabiliza vazio/0)
#      * Matrícula Ativa: Coluna H deve conter "MA"
#      * Inclusão: Coluna N == "Sim" (case/trim)
#      * Plano de Ação: Coluna P com profissional válido,
#                       MAS SOMENTE se o aluno também for Inclusão
#      * Profissionais: únicos por turma (dedupe case-insensitive e colapsando espaços)
#  - Alertas:
#      * Turmas com 2+ profissionais distintos (Plano de Ação)
#      * Aluno com profissional na P, mas sem Inclusão na N
#  - Preenchimento por mapeamento automático do MODELO
# ==========================================================

# OBS: este código assume que você já tem:
# - app = Flask(__name__)
# - app.config["UPLOAD_FOLDER"]
# no seu arquivo principal.


@app.route("/quantinclusao", methods=["GET", "POST"])
@login_required
def quantinclusao():
    if request.method == "POST":
        reg_file = (
            request.files.get("lista_regular")
            or request.files.get("lista_fundamental")
            or request.files.get("lista")
        )
        responsavel = request.form.get("responsavel")

        if not reg_file or reg_file.filename == "":
            flash("Selecione o arquivo da Lista Piloto (Regular/Fundamental).", "error")
            return redirect(url_for("quantinclusao"))

        valid, message = validate_excel_upload(reg_file)
        if not valid:
            flash(message, "error")
            return redirect(url_for("quantinclusao"))

        if not responsavel or responsavel.strip() == "":
            flash("Informe o Responsável pelo preenchimento.", "error")
            return redirect(url_for("quantinclusao"))

        reg_path = save_excel_upload(
            reg_file,
            app.config["UPLOAD_FOLDER"],
            prefix="regular",
        )

        model_path = os.path.join("modelos", "Quadro Quantitativo de Inclusão - Modelo.xlsx")
        try:
            result = build_quant_inclusao_file(
                reg_path,
                model_path,
                responsavel.strip(),
            )
        except QuantInclusaoError as e:
            flash(str(e), "error")
            return redirect(url_for("quantinclusao"))

        resp = send_file(
            result.output,
            as_attachment=True,
            download_name=result.filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Dados pessoais ficam na aba ALERTAS do Excel, não em headers HTTP.
        if result.alerts:
            resp.headers["X-QuantInclusao-Alerts-Count"] = str(len(result.alerts))

        if result.plan_without_inclusion_alerts:
            resp.headers["X-QuantInclusao-PlanWithoutInclusion-Count"] = str(
                len(result.plan_without_inclusion_alerts)
            )

        return resp

    return render_template("quantinclusao.html")

def build_deadline_alerts(today=None):
    return build_deadline_alerts_service(
        today=today,
        holidays_path=app.config.get("HOLIDAYS_JSON_PATH"),
        weekly_due_weekday=app.config.get("INFORMATIVO_WEEKDAY_DUE", 4),
    )


# Injeta automaticamente para TODOS os templates
@app.context_processor
def _inject_deadline_alerts():
    return {"deadline_alerts": build_deadline_alerts()}

# ==========================================================
#  MAIN
# ==========================================================

if __name__ == "__main__":
    app.run(debug=True)
