import calendar
import json
import os
from dataclasses import dataclass
from datetime import date, datetime, timedelta


try:
    from zoneinfo import ZoneInfo

    _TZ = ZoneInfo("America/Sao_Paulo")
except Exception:
    _TZ = None


_HOLIDAYS_CACHE = {}


@dataclass
class DeadlineAlert:
    key: str
    title: str
    due_base: date
    due_adjusted: date
    window_before: int
    window_after: int
    message: str


def clear_holidays_cache():
    _HOLIDAYS_CACHE.clear()


def today_sao_paulo() -> date:
    if _TZ:
        return datetime.now(_TZ).date()
    return date.today()


def load_holidays(path: str | None) -> dict:
    cache_key = os.path.abspath(path) if path else ""
    if cache_key in _HOLIDAYS_CACHE:
        return _HOLIDAYS_CACHE[cache_key]

    holidays = {
        "loaded": True,
        "dates": set(),
        "names": {},
        "error": None,
        "path": path,
    }

    try:
        if not path or not os.path.exists(path):
            holidays["error"] = f"feriados.json nao encontrado em: {path}"
        else:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            if not isinstance(data, dict):
                holidays["error"] = "feriados.json invalido: esperado um objeto JSON (dict)."
            else:
                for key, value in data.items():
                    if not isinstance(key, str):
                        continue
                    try:
                        dt = date(int(key[0:4]), int(key[5:7]), int(key[8:10]))
                    except Exception:
                        continue

                    holidays["dates"].add(dt)
                    name = str(value).strip() if value is not None else ""
                    if name:
                        holidays["names"][dt] = name
    except Exception as exc:
        holidays["error"] = str(exc)

    _HOLIDAYS_CACHE[cache_key] = holidays
    return holidays


def holiday_name(d: date, holidays: dict) -> str:
    return holidays["names"].get(d, "")


def is_business_day(d: date, holidays: dict) -> bool:
    if d.weekday() >= 5:
        return False
    return d not in holidays["dates"]


def next_business_day(d: date, holidays: dict) -> date:
    cur = d
    while not is_business_day(cur, holidays):
        cur += timedelta(days=1)
    return cur


def previous_business_day(d: date, holidays: dict) -> date:
    cur = d
    while not is_business_day(cur, holidays):
        cur -= timedelta(days=1)
    return cur


def last_day_of_month(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def add_months(base: date, months: int) -> date:
    year = base.year + (base.month - 1 + months) // 12
    month = (base.month - 1 + months) % 12 + 1
    day = min(base.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def fmt_br(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def window_hit(today: date, due: date, before: int, after: int) -> bool:
    return (due - timedelta(days=before)) <= today <= (due + timedelta(days=after))


def compute_due_day20(ref: date, holidays: dict) -> tuple[date, date]:
    base = date(ref.year, ref.month, 20)
    adjusted = next_business_day(base, holidays)

    if ref > (adjusted + timedelta(days=2)):
        next_base = add_months(base, 1)
        next_base = date(next_base.year, next_base.month, 20)
        return next_base, next_business_day(next_base, holidays)

    return base, adjusted


def compute_due_month_end(ref: date, holidays: dict) -> tuple[date, date]:
    base = last_day_of_month(ref.year, ref.month)
    adjusted = base if is_business_day(base, holidays) else previous_business_day(base, holidays)

    if ref > (adjusted + timedelta(days=2)):
        next_month = add_months(date(ref.year, ref.month, 1), 1)
        next_base = last_day_of_month(next_month.year, next_month.month)
        next_adjusted = (
            next_base if is_business_day(next_base, holidays) else previous_business_day(next_base, holidays)
        )
        return next_base, next_adjusted

    return base, adjusted


def compute_due_weekly(ref: date, holidays: dict, weekly_due_weekday: int = 4) -> tuple[date, date]:
    due_weekday = max(0, min(6, int(weekly_due_weekday)))
    monday = ref - timedelta(days=ref.weekday())
    base = monday + timedelta(days=due_weekday)
    adjusted = next_business_day(base, holidays)

    if ref > (adjusted + timedelta(days=1)):
        next_monday = monday + timedelta(days=7)
        next_base = next_monday + timedelta(days=due_weekday)
        return next_base, next_business_day(next_base, holidays)

    return base, adjusted


def _append_status(alerts: list[DeadlineAlert], today: date) -> list[dict]:
    output = []
    for alert in alerts:
        delta = (alert.due_adjusted - today).days
        if delta > 0:
            status = f"Faltam {delta} dia(s)."
        elif delta == 0:
            status = "Vence hoje."
        else:
            status = f"Vencido há {abs(delta)} dia(s)."

        output.append(
            {
                "key": alert.key,
                "title": alert.title,
                "message": alert.message,
                "due": fmt_br(alert.due_adjusted),
                "status": status,
            }
        )

    return output


def build_deadline_alerts(
    today: date | None = None,
    *,
    holidays_path: str | None = None,
    weekly_due_weekday: int = 4,
) -> list[dict]:
    today = today or today_sao_paulo()
    holidays = load_holidays(holidays_path)

    day20_base, day20_adjusted = compute_due_day20(today, holidays)
    month_end_base, month_end_adjusted = compute_due_month_end(today, holidays)
    weekly_base, weekly_adjusted = compute_due_weekly(today, holidays, weekly_due_weekday)

    alerts: list[DeadlineAlert] = []

    if window_hit(today, day20_adjusted, before=2, after=2):
        extra = ""
        if day20_adjusted != day20_base:
            name = holiday_name(day20_base, holidays)
            extra = f" (ajustado para próximo dia útil{': ' + name if name else ''})"
        alerts.append(
            DeadlineAlert(
                key="quant_inclusao",
                title="Prazo: Quantitativo de Inclusão",
                due_base=day20_base,
                due_adjusted=day20_adjusted,
                window_before=2,
                window_after=2,
                message=f"O Quadro Quantitativo de Inclusão deve ser enviado até {fmt_br(day20_adjusted)}{extra}.",
            )
        )

    if window_hit(today, month_end_adjusted, before=2, after=2):
        extra = ""
        if month_end_adjusted != month_end_base:
            name = holiday_name(month_end_base, holidays)
            extra = f" (ajustado para último dia útil do mês{': ' + name if name else ''})"
        alerts.append(
            DeadlineAlert(
                key="atendimento_mensal",
                title="Prazo: Atendimento Mensal",
                due_base=month_end_base,
                due_adjusted=month_end_adjusted,
                window_before=2,
                window_after=2,
                message=f"O Quadro de Atendimento Mensal deve ser enviado até {fmt_br(month_end_adjusted)}{extra}.",
            )
        )

        alerts.append(
            DeadlineAlert(
                key="quant_mensal_te",
                title="Prazo: Quantitativo Mensal de Transferências Expedidas",
                due_base=month_end_base,
                due_adjusted=month_end_adjusted,
                window_before=2,
                window_after=2,
                message=(
                    "O Quadro Mensal de Transferências Expedidas deve ser enviado "
                    f"até {fmt_br(month_end_adjusted)}{extra}."
                ),
            )
        )

    if window_hit(today, weekly_adjusted, before=1, after=3):
        extra = ""
        if weekly_adjusted != weekly_base:
            name = holiday_name(weekly_base, holidays)
            extra = f" (ajustado para próximo dia útil{': ' + name if name else ''})"
        alerts.append(
            DeadlineAlert(
                key="informativo_semanal",
                title="Prazo: Informativo Semanal",
                due_base=weekly_base,
                due_adjusted=weekly_adjusted,
                window_before=1,
                window_after=1,
                message=f"O Informativo Semanal deve ser enviado até {fmt_br(weekly_adjusted)}{extra}.",
            )
        )

    return _append_status(alerts, today)
