import json
import tempfile
import unittest
from datetime import date
from pathlib import Path

from services.prazos import build_deadline_alerts, clear_holidays_cache


def write_holidays(data):
    tmp = tempfile.TemporaryDirectory()
    path = Path(tmp.name) / "feriados.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    return tmp, str(path)


class DeadlineAlertTests(unittest.TestCase):
    def tearDown(self):
        clear_holidays_cache()

    def test_day20_moves_to_next_business_day_when_holiday(self):
        tmp, path = write_holidays({"2026-02-20": "Feriado ficticio"})
        self.addCleanup(tmp.cleanup)

        alerts = build_deadline_alerts(date(2026, 2, 23), holidays_path=path)
        by_key = {alert["key"]: alert for alert in alerts}

        self.assertEqual(by_key["quant_inclusao"]["due"], "23/02/2026")
        self.assertIn("Feriado ficticio", by_key["quant_inclusao"]["message"])

    def test_month_end_uses_previous_business_day(self):
        tmp, path = write_holidays({})
        self.addCleanup(tmp.cleanup)

        alerts = build_deadline_alerts(date(2026, 5, 29), holidays_path=path)
        by_key = {alert["key"]: alert for alert in alerts}

        self.assertEqual(by_key["atendimento_mensal"]["due"], "29/05/2026")
        self.assertEqual(by_key["quant_mensal_te"]["due"], "29/05/2026")

    def test_weekly_alert_moves_to_next_business_day_when_holiday(self):
        tmp, path = write_holidays({"2026-04-03": "Feriado semanal ficticio"})
        self.addCleanup(tmp.cleanup)

        alerts = build_deadline_alerts(date(2026, 4, 5), holidays_path=path)
        by_key = {alert["key"]: alert for alert in alerts}

        self.assertEqual(by_key["informativo_semanal"]["due"], "06/04/2026")
        self.assertIn("Feriado semanal ficticio", by_key["informativo_semanal"]["message"])


if __name__ == "__main__":
    unittest.main()
