import os
import secrets
from datetime import datetime

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None


DEFAULT_UPLOAD_FOLDER = "uploads"
DEFAULT_MAX_CONTENT_LENGTH_MB = 50
DEFAULT_SESSION_COOKIE_SAMESITE = "Lax"


def load_environment():
    if load_dotenv:
        load_dotenv()


def env_bool(name, default=False):
    raw = os.getenv(name)
    if raw is None:
        return default
    return str(raw).strip().lower() in {"1", "true", "t", "yes", "y", "on"}


def env_int(name, default):
    raw = os.getenv(name)
    if raw is None or str(raw).strip() == "":
        return default
    try:
        return int(raw)
    except ValueError:
        print(f"[AVISO] Valor invalido para {name}: {raw!r}. Usando {default}.")
        return default


def resolve_secret_key():
    key = os.getenv("FLASK_SECRET_KEY")
    if key:
        return key
    print(
        "[AVISO] FLASK_SECRET_KEY nao configurada. "
        "Usando chave temporaria; sessoes serao invalidadas ao reiniciar."
    )
    return secrets.token_urlsafe(32)


def holidays_json_path(app):
    return os.path.join(app.root_path, "modelos", "feriados.json")


def configure_app(app):
    load_environment()

    app.secret_key = resolve_secret_key()
    access_token = os.getenv("ACCESS_TOKEN")
    if not access_token:
        print("[AVISO] ACCESS_TOKEN nao configurado. Login ficara indisponivel ate configurar a variavel.")

    confere_password = os.getenv("CONFERE_PASSWORD")
    if not confere_password:
        print(
            "[AVISO] CONFERE_PASSWORD nao configurado. "
            "Acesso externo ao conferidor ficara indisponivel ate configurar a variavel."
        )

    school_year = env_int("SCHOOL_YEAR", datetime.now().year)

    app.config["UPLOAD_FOLDER"] = DEFAULT_UPLOAD_FOLDER
    app.config["MAX_CONTENT_LENGTH"] = (
        env_int("MAX_CONTENT_LENGTH_MB", DEFAULT_MAX_CONTENT_LENGTH_MB) * 1024 * 1024
    )
    app.config["SESSION_COOKIE_HTTPONLY"] = True
    app.config["SESSION_COOKIE_SAMESITE"] = DEFAULT_SESSION_COOKIE_SAMESITE
    app.config["SESSION_COOKIE_SECURE"] = env_bool("SESSION_COOKIE_SECURE", False)
    app.config["SCHOOL_YEAR"] = school_year
    app.config["CONCLUSAO_5ANO_DATE_TEXT"] = os.getenv(
        "CONCLUSAO_5ANO_DATE_TEXT",
        f"Praia Grande, 22 de dezembro de {school_year}",
    )
    app.config["HOLIDAYS_JSON_PATH"] = holidays_json_path(app)
    app.config["INFORMATIVO_WEEKDAY_DUE"] = env_int("INFORMATIVO_WEEKDAY_DUE", 4)
    app.config["CONFERE_PASSWORD"] = confere_password

    return access_token
