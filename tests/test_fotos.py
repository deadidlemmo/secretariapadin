import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from PIL import Image

from services.fotos import get_student_photo_url, save_student_photo, student_has_photo


def make_png_1x1():
    stream = BytesIO()
    Image.new("RGB", (1, 1), color="white").save(stream, format="PNG")
    return stream.getvalue()


PNG_1X1 = make_png_1x1()


class DummyImageUpload:
    def __init__(self, filename, content=PNG_1X1):
        self.filename = filename
        self.stream = BytesIO(content)

    def save(self, path):
        self.stream.seek(0)
        Path(path).write_bytes(self.stream.read())


class FotosServiceTests(unittest.TestCase):
    def test_save_student_photo_uses_rm_filename_and_returns_url(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            saved = save_student_photo(DummyImageUpload("foto.PNG"), "123", photos_dir=tmpdir)

            self.assertEqual(saved["filename"], "123.png")
            self.assertEqual(saved["url"], "/static/fotos/123.png")
            self.assertTrue(Path(saved["path"]).exists())

    def test_find_student_photo_url_uses_existing_photo(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            Path(tmpdir, "456.jpg").write_bytes(b"fake")

            self.assertTrue(student_has_photo(456, photos_dir=tmpdir))
            self.assertEqual(get_student_photo_url(456, photos_dir=tmpdir), "/static/fotos/456.jpg")

    def test_save_student_photo_rejects_missing_rm(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(ValueError):
                save_student_photo(DummyImageUpload("foto.png"), "", photos_dir=tmpdir)


if __name__ == "__main__":
    unittest.main()
