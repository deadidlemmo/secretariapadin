import json
import tempfile
import unittest
from pathlib import Path

from services.carteirinhas_log import get_printed_set, load_print_log, mark_printed_rms, save_print_log


class CarteirinhasLogTests(unittest.TestCase):
    def test_migrates_legacy_json_to_sqlite(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            json_path = Path(tmpdir) / "carteirinhas_print_log.json"
            db_path = Path(tmpdir) / "carteirinhas_print_log.sqlite3"
            json_path.write_text(
                json.dumps(
                    {
                        "2026": {
                            "printed_rms": [101, "102", "invalido"],
                            "printed_at": {"101": "2026-01-10T10:00:00"},
                        }
                    }
                ),
                encoding="utf-8",
            )

            self.assertEqual(get_printed_set(2026, db_path=str(db_path), json_path=str(json_path)), {101, 102})

    def test_mark_printed_rms_counts_only_new_records(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = str(Path(tmpdir) / "carteirinhas_print_log.sqlite3")

            added, total = mark_printed_rms(2026, [101, 102], db_path=db_path, json_path="")
            self.assertEqual((added, total), (2, 2))

            added, total = mark_printed_rms(2026, [102, 103], db_path=db_path, json_path="")
            self.assertEqual((added, total), (1, 3))
            self.assertEqual(get_printed_set(2026, db_path=db_path, json_path=""), {101, 102, 103})

    def test_save_and_load_print_log_shape(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = str(Path(tmpdir) / "carteirinhas_print_log.sqlite3")
            data = {
                "2026": {
                    "printed_rms": [201],
                    "printed_at": {"201": "2026-02-01T09:00:00"},
                }
            }

            save_print_log(data, db_path=db_path, json_path="")
            loaded = load_print_log(db_path=db_path, json_path="")

            self.assertEqual(loaded["2026"]["printed_rms"], [201])
            self.assertEqual(loaded["2026"]["printed_at"]["201"], "2026-02-01T09:00:00")


if __name__ == "__main__":
    unittest.main()
