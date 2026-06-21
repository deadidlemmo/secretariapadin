import os
import uuid

from werkzeug.utils import secure_filename

try:
    from PIL import Image, UnidentifiedImageError
except Exception:
    Image = None
    UnidentifiedImageError = Exception


ALLOWED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}
ALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}


def allowed_extension(filename, allowed_extensions):
    _, ext = os.path.splitext(filename or "")
    return ext.lower() in allowed_extensions


def allowed_image_file(filename):
    return allowed_extension(filename, ALLOWED_IMAGE_EXTENSIONS)


def allowed_excel_file(filename):
    return allowed_extension(filename, ALLOWED_EXCEL_EXTENSIONS)


def unique_secure_filename(filename, prefix=None):
    safe_name = secure_filename(filename or "")
    if not safe_name:
        return ""
    parts = [part for part in [prefix, uuid.uuid4().hex, safe_name] if part]
    return "_".join(parts)


def save_upload(file_storage, upload_folder, *, prefix=None):
    filename = unique_secure_filename(file_storage.filename or "", prefix=prefix)
    if not filename:
        raise ValueError("Nenhum arquivo selecionado.")
    os.makedirs(upload_folder, exist_ok=True)
    file_path = os.path.join(upload_folder, filename)
    file_storage.save(file_path)
    return file_path


def save_excel_upload(file_storage, upload_folder, *, prefix=None, required=True):
    valid, message = validate_excel_upload(file_storage, required=required)
    if not valid:
        raise ValueError(message)
    if not file_storage or file_storage.filename == "":
        return None
    return save_upload(file_storage, upload_folder, prefix=prefix)


def validate_image_upload(file_storage):
    if not file_storage or file_storage.filename == "":
        return False, "Nenhuma foto selecionada."
    if not allowed_image_file(file_storage.filename):
        return False, "Formato de imagem nao permitido. Envie JPG, PNG, GIF ou BMP."
    if Image is None:
        return True, None

    try:
        stream_pos = file_storage.stream.tell()
    except Exception:
        stream_pos = None

    try:
        with Image.open(file_storage.stream) as img:
            img.verify()
    except (UnidentifiedImageError, OSError, ValueError, SyntaxError):
        return False, "Arquivo de imagem invalido ou corrompido."
    finally:
        try:
            file_storage.stream.seek(stream_pos or 0)
        except Exception:
            pass

    return True, None


def validate_excel_upload(file_storage, *, required=True):
    if not file_storage or file_storage.filename == "":
        if required:
            return False, "Nenhum arquivo selecionado."
        return True, None
    if not allowed_excel_file(file_storage.filename):
        return False, "Formato invalido. Envie uma planilha .xlsx, .xls ou .xlsm."
    return True, None
