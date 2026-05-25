import os

from werkzeug.utils import secure_filename

from utils.uploads import validate_image_upload


PHOTO_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".gif")
DEFAULT_PHOTOS_DIR = os.path.join("static", "fotos")
DEFAULT_URL_PREFIX = "/static/fotos"


def normalize_rm(rm) -> str:
    return str(rm or "").strip()


def photo_filename(rm, extension: str) -> str:
    rm_text = normalize_rm(rm)
    return secure_filename(f"{rm_text}{extension.lower()}")


def photo_path(rm, extension: str, photos_dir=DEFAULT_PHOTOS_DIR) -> str:
    return os.path.join(photos_dir, photo_filename(rm, extension))


def find_student_photo(rm, photos_dir=DEFAULT_PHOTOS_DIR, url_prefix=DEFAULT_URL_PREFIX):
    for extension in PHOTO_EXTENSIONS:
        filename = photo_filename(rm, extension)
        path = os.path.join(photos_dir, filename)
        if os.path.exists(path):
            return {
                "filename": filename,
                "path": path,
                "url": f"{url_prefix}/{filename}",
            }
    return None


def get_student_photo_url(rm, photos_dir=DEFAULT_PHOTOS_DIR, url_prefix=DEFAULT_URL_PREFIX):
    found = find_student_photo(rm, photos_dir=photos_dir, url_prefix=url_prefix)
    return found["url"] if found else None


def student_has_photo(rm, photos_dir=DEFAULT_PHOTOS_DIR) -> bool:
    return find_student_photo(rm, photos_dir=photos_dir) is not None


def save_student_photo(file_storage, rm, photos_dir=DEFAULT_PHOTOS_DIR, url_prefix=DEFAULT_URL_PREFIX):
    rm_text = normalize_rm(rm)
    if not rm_text:
        raise ValueError("RM nao fornecido.")

    valid, message = validate_image_upload(file_storage)
    if not valid:
        raise ValueError(message)

    safe_original = secure_filename(file_storage.filename or "")
    _, extension = os.path.splitext(safe_original)
    filename = photo_filename(rm_text, extension)
    if not filename:
        raise ValueError("Nome de arquivo invalido.")

    os.makedirs(photos_dir, exist_ok=True)
    path = os.path.join(photos_dir, filename)
    file_storage.save(path)

    return {
        "filename": filename,
        "path": path,
        "url": f"{url_prefix}/{filename}",
    }
