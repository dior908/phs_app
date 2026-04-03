import json
import os
import sqlite3
from datetime import datetime
from typing import Optional

# Путь к файлу конфига команды
TEAM_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "team_config.json")


def _load_config() -> dict:
    if not os.path.exists(TEAM_CONFIG_PATH):
        return {"inactive": []}
    with open(TEAM_CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_config(config: dict):
    with open(TEAM_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def get_inactive_ids() -> list:
    """Возвращает список telegram_id уволенных сотрудников."""
    return _load_config().get("inactive", [])


def deactivate_employee(telegram_id: int):
    """Пометить сотрудника как уволенного."""
    config = _load_config()
    if telegram_id not in config["inactive"]:
        config["inactive"].append(telegram_id)
    _save_config(config)


def activate_employee(telegram_id: int):
    """Восстановить сотрудника."""
    config = _load_config()
    config["inactive"] = [i for i in config["inactive"] if i != telegram_id]
    _save_config(config)


def is_active(telegram_id: int) -> bool:
    return telegram_id not in get_inactive_ids()


def get_active_users(users_db: str) -> list:
    """Возвращает только активных сотрудников."""
    inactive = get_inactive_ids()
    with sqlite3.connect(users_db) as conn:
        cur = conn.cursor()
        cur.execute("SELECT telegram_id, full_name, region, phone_number FROM users ORDER BY full_name")
        all_users = cur.fetchall()
    return [u for u in all_users if u[0] not in inactive]


def get_all_users_with_status(users_db: str) -> list:
    """Возвращает всех сотрудников с полем статуса (active: bool)."""
    inactive = get_inactive_ids()
    with sqlite3.connect(users_db) as conn:
        cur = conn.cursor()
        cur.execute("SELECT telegram_id, full_name, region, phone_number FROM users ORDER BY full_name")
        all_users = cur.fetchall()
    return [(u[0], u[1], u[2], u[3], u[0] not in inactive) for u in all_users]


def get_first_visit_date(visits_db: str, telegram_id: int) -> Optional[str]:
    """Дата первого визита сотрудника (как дата начала работы)."""
    with sqlite3.connect(visits_db) as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT date_time FROM visits WHERE telegram_id = ? ORDER BY date_time ASC LIMIT 1",
            (telegram_id,)
        )
        row = cur.fetchone()

    if row:
        try:
            dt = datetime.strptime(row[0], "%d.%m.%Y %H:%M")
            return dt.strftime("%d.%m.%Y")
        except Exception:
            return row[0]

    return None


def get_days_worked(visits_db: str, telegram_id: int) -> int:
    """Количество дней от первого визита до сегодня."""
    first = get_first_visit_date(visits_db, telegram_id)
    if not first:
        return 0
    try:
        start = datetime.strptime(first, "%d.%m.%Y")
        return (datetime.now() - start).days + 1
    except Exception:
        return 0
