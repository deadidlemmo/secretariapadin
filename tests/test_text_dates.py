import unittest
from datetime import datetime

import pandas as pd

from utils.dates import (
    data_extenso,
    detect_te_date_from_obs_flexible,
    extract_te_date_from_text,
    parse_date_flexible,
    parse_period_date,
    parse_user_date,
)
from utils.text import build_colmap, find_df_col, is_missing_text, is_missing_value, norm_header_compact, pick_col


class TextUtilsTests(unittest.TestCase):
    def test_header_normalization_and_column_lookup(self):
        df = pd.DataFrame([["5ºA", "Rede", "Aluno Ficticio"]], columns=["SÉRIE", "Local TE", "Nome"])
        colmap = build_colmap(df)

        self.assertEqual(norm_header_compact("Série"), "SERIE")
        self.assertEqual(pick_col(colmap, "LOCAL_TE", "LOCAL TE"), "Local TE")
        self.assertEqual(find_df_col(df, ["serie"]), "SÉRIE")

    def test_missing_value_rules(self):
        for value in ("", "0", "-", "nan", None):
            self.assertTrue(is_missing_value(value))
            self.assertTrue(is_missing_text(value))
        self.assertFalse(is_missing_value("Dentro da Rede"))


class DateUtilsTests(unittest.TestCase):
    def test_date_parsers_accept_current_formats(self):
        self.assertEqual(parse_user_date("31/01/2026"), datetime(2026, 1, 31))
        self.assertEqual(parse_period_date("2026-01-31", "a data final"), datetime(2026, 1, 31))
        self.assertEqual(parse_date_flexible("16/01", default_year=2026), datetime(2026, 1, 16))
        self.assertEqual(parse_date_flexible("16/01/26"), datetime(2026, 1, 16))

    def test_data_extenso(self):
        self.assertEqual(data_extenso(datetime(2026, 3, 2)), "2 de março de 2026")

    def test_te_date_detection(self):
        dt, match, inferred = extract_te_date_from_text(
            "Aluno TE - 16/01",
            datetime(2026, 1, 1),
            datetime(2026, 1, 31),
        )
        self.assertEqual(dt, datetime(2026, 1, 16))
        self.assertEqual(match, "TE - 16/01")
        self.assertTrue(inferred)

        dt, rule, match, inferred = detect_te_date_from_obs_flexible("TE - 16/01/26")
        self.assertEqual(dt, datetime(2026, 1, 16))
        self.assertEqual(rule, "OBS:TE_DATE")
        self.assertEqual(match, "TE - 16/01/26")
        self.assertFalse(inferred)


if __name__ == "__main__":
    unittest.main()
