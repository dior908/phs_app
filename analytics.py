import os
import sqlite3
from collections import defaultdict
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

import pytz

# ─── КОНСТАНТЫ НОРМ ───────────────────────────────────────────────────────────
MIN_DOCTORS_DAY  = 8
MAX_DOCTORS_DAY  = 12
MIN_PHARM_DAY    = 6
MAX_PHARM_DAY    = 10

TZ = pytz.timezone("Asia/Tashkent")

DAYS_RU = {0: "Пн", 1: "Вт", 2: "Ср", 3: "Чт", 4: "Пт", 5: "Сб", 6: "Вс"}

# ─── СТИЛИ ────────────────────────────────────────────────────────────────────
def _hdr_font():  return Font(name="Arial", bold=True, color="FFFFFF", size=11)
def _body_font(): return Font(name="Arial", size=10)
def _title_font():return Font(name="Arial", bold=True, size=13, color="1F4E79")

def _fill(hex_color): return PatternFill("solid", start_color=hex_color)

FILL_HEADER   = _fill("1F4E79")
FILL_SUBHEAD  = _fill("2E75B6")
FILL_GREEN    = _fill("E2EFDA")
FILL_YELLOW   = _fill("FFEB9C")
FILL_ORANGE   = _fill("FCE4D6")
FILL_RED      = _fill("FF7676")
FILL_BLUE     = _fill("DEEAF1")
FILL_PURPLE   = _fill("E2D0F0")
FILL_ALT      = _fill("F5F5F5")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

def _border():
    thin = Side(style="thin", color="BBBBBB")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def _apply_header(cell, fill=None):
    cell.font      = _hdr_font()
    cell.fill      = fill or FILL_HEADER
    cell.alignment = CENTER
    cell.border    = _border()

def _apply_body(cell, fill=None, align=CENTER):
    cell.font      = _body_font()
    cell.fill      = fill or _fill("FFFFFF")
    cell.alignment = align
    cell.border    = _border()

def _set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ─── ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ──────────────────────────────────────────────────

def _working_days(start_dt, end_dt):
    return sum(
        1 for d in range((end_dt - start_dt).days + 1)
        if (start_dt + timedelta(days=d)).weekday() < 5
    )

def _norms(working_days):
    return {
        "min_doc":   MIN_DOCTORS_DAY * working_days,
        "max_doc":   MAX_DOCTORS_DAY * working_days,
        "min_ph":    MIN_PHARM_DAY   * working_days,
        "max_ph":    MAX_PHARM_DAY   * working_days,
        "min_total": (MIN_DOCTORS_DAY + MIN_PHARM_DAY) * working_days,
        "max_total": (MAX_DOCTORS_DAY + MAX_PHARM_DAY) * working_days,
    }

def _verdict(total, min_total, max_total):
    if total < min_total:              return "❌ Норма не выполнена"
    elif total == min_total:           return "✅ Выполнен минимум"
    elif min_total < total < max_total:return "👍 Выполнено средне"
    elif total == max_total:           return "🏆 Выполнен максимум"
    else:                              return "🚀 Выше максимума"

def _verdict_fill(verdict):
    if "не выполнена" in verdict: return FILL_RED
    if "минимум"      in verdict: return FILL_YELLOW
    if "средне"       in verdict: return FILL_ORANGE
    if "максимум"     in verdict: return FILL_GREEN
    return FILL_PURPLE  # выше максимума

def _pct(val, norm):
    return round(val / norm * 100, 1) if norm else 0

def _parse_dt(dt_str):
    try:
        return datetime.strptime(dt_str, "%d.%m.%Y %H:%M")
    except Exception:
        return None

# ─── ЧТЕНИЕ ДАННЫХ ────────────────────────────────────────────────────────────

def fetch_visits_range(visits_db, start_dt, end_dt):
    """
    Фильтрация в Python, а не в SQLite — потому что дата хранится как
    текст в формате ДД.ММ.ГГГГ ЧЧ:ММ, и SQLite BETWEEN сравнивает
    строки лексикографически (неверно для этого формата).
    """
    # Убираем timezone для корректного сравнения
    if hasattr(start_dt, 'tzinfo') and start_dt.tzinfo:
        start_dt = start_dt.replace(tzinfo=None)
    if hasattr(end_dt, 'tzinfo') and end_dt.tzinfo:
        end_dt = end_dt.replace(tzinfo=None)

    with sqlite3.connect(visits_db) as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM visits ORDER BY telegram_id, date_time")
        all_rows = cur.fetchall()

    result = []
    for row in all_rows:
        dt = _parse_dt(row[4])  # date_time — индекс 4
        if dt and start_dt <= dt <= end_dt:
            result.append(row)
    return result

def fetch_all_users(users_db):
    with sqlite3.connect(users_db) as conn:
        cur = conn.cursor()
        cur.execute("SELECT telegram_id, full_name, region FROM users ORDER BY full_name")
        return cur.fetchall()

# ─── ЛИСТ 1: ОСНОВНАЯ ТАБЛИЦА ─────────────────────────────────────────────────

def _sheet_main(wb, users, visits_by_user, working_days, period_label, start_dt, end_dt):
    ws = wb.active
    ws.title = "📊 Аналитика"
    norms = _norms(working_days)

    # Заголовок
    ws.merge_cells("A1:N1")
    ws["A1"] = f"Аналитика команды | {period_label} | {start_dt.strftime('%d.%m.%Y')} — {end_dt.strftime('%d.%m.%Y')}"
    ws["A1"].font      = _title_font()
    ws["A1"].alignment = CENTER
    ws["A1"].fill      = FILL_BLUE
    ws.row_dimensions[1].height = 28

    # Нормы-справка
    ws.merge_cells("A2:N2")
    ws["A2"] = (f"Нормы: Врачи {MIN_DOCTORS_DAY}–{MAX_DOCTORS_DAY}/день  |  "
                f"Аптеки/Оптом {MIN_PHARM_DAY}–{MAX_PHARM_DAY}/день  |  "
                f"Рабочих дней в периоде: {working_days}")
    ws["A2"].font      = Font(name="Arial", size=10, italic=True, color="555555")
    ws["A2"].alignment = LEFT
    ws.row_dimensions[2].height = 18

    headers = [
        "№", "Сотрудник", "Регион",
        "Врачи", f"Норма (мин {norms['min_doc']})", "% врачи",
        "Аптеки", f"Норма (мин {norms['min_ph']})", "% аптеки",
        "Оптом", "Итого визитов", f"Норма итого (мин {norms['min_total']})",
        "% итого", "Вердикт"
    ]
    ws.append(headers)
    for col, cell in enumerate(ws[3], 1):
        _apply_header(cell)
    ws.row_dimensions[3].height = 40

    user_map = {u[0]: u for u in users}
    summary_rows = []

    for idx, user in enumerate(users, 1):
        uid    = user[0]
        visits = visits_by_user.get(uid, [])
        absent = len([v for v in visits if v[3] == "Не вышел"])
        wd     = max(working_days - absent, 1)
        n      = _norms(wd)

        doctors = len([v for v in visits if v[3] == "🩺 Врач"])
        pharm   = len([v for v in visits if v[3] == "💊 Аптека"])
        dist    = len([v for v in visits if v[3] == "🚚 Дистрибьютор"])
        total   = doctors + pharm + dist
        verdict = _verdict(total, n["min_total"], n["max_total"])

        row = [
            idx, user[1], user[2],
            doctors, n["min_doc"], f"{_pct(doctors, n['min_doc'])}%",
            pharm,   n["min_ph"],  f"{_pct(pharm, n['min_ph'])}%",
            dist,
            total,   n["min_total"], f"{_pct(total, n['min_total'])}%",
            verdict
        ]
        ws.append(row)
        summary_rows.append((user[1], doctors, pharm, dist, total, verdict))

        vfill = _verdict_fill(verdict)
        alt   = idx % 2 == 0
        row_fill = FILL_ALT if alt else _fill("FFFFFF")

        for col_i, cell in enumerate(ws[idx + 3], 1):
            if col_i == 14:
                _apply_body(cell, vfill, LEFT)
            elif col_i in (6, 9, 13):
                pct_val = float(str(cell.value).replace("%","")) if cell.value else 0
                pf = FILL_GREEN if pct_val >= 100 else FILL_YELLOW if pct_val >= 60 else FILL_RED
                _apply_body(cell, pf)
            else:
                _apply_body(cell, row_fill)

    ws.row_dimensions[3].height = 40
    ws.freeze_panes = "A4"
    _set_col_widths(ws, [4, 22, 14, 8, 16, 10, 8, 16, 10, 8, 12, 20, 10, 26])

    return summary_rows

# ─── ЛИСТ 2: АКТИВНОСТЬ ПО ДНЯМ НЕДЕЛИ ───────────────────────────────────────

def _sheet_weekdays(wb, visits_by_user, users):
    ws = wb.create_sheet("📅 По дням недели")
    user_map = {u[0]: u[1] for u in users}

    ws.merge_cells("A1:H1")
    ws["A1"] = "Активность по дням недели"
    ws["A1"].font = _title_font(); ws["A1"].alignment = CENTER; ws["A1"].fill = FILL_BLUE

    headers = ["Сотрудник", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Итого"]
    ws.append(headers)
    for cell in ws[2]:
        _apply_header(cell)

    for idx, (uid, visits) in enumerate(visits_by_user.items(), 3):
        day_counts = defaultdict(int)
        for v in visits:
            if v[3] == "Не вышел": continue
            dt = _parse_dt(v[4])
            if dt: day_counts[dt.weekday()] += 1

        row = [user_map.get(uid, str(uid))]
        for d in range(6):
            row.append(day_counts.get(d, 0))
        row.append(sum(day_counts.values()))
        ws.append(row)

        fill = FILL_ALT if idx % 2 == 0 else _fill("FFFFFF")
        for cell in ws[idx]:
            _apply_body(cell, fill)

    _set_col_widths(ws, [22, 6, 6, 6, 6, 6, 6, 8])
    ws.freeze_panes = "A3"

    # График
    if len(visits_by_user) > 0:
        chart = BarChart()
        chart.type    = "col"
        chart.title   = "Визиты по дням недели"
        chart.y_axis.title = "Кол-во визитов"
        chart.style   = 10
        chart.width   = 20
        chart.height  = 14

        data = Reference(ws, min_col=2, max_col=7,
                         min_row=2, max_row=2 + len(visits_by_user))
        cats = Reference(ws, min_col=1, min_row=3, max_row=2 + len(visits_by_user))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, "J2")

# ─── ЛИСТ 3: ТОП И АУТСАЙДЕРЫ ОРГАНИЗАЦИЙ ────────────────────────────────────

def _sheet_orgs(wb, all_visits):
    ws = wb.create_sheet("🏆 Орг-ции")

    org_counts = defaultdict(int)
    org_cats   = defaultdict(set)
    for v in all_visits:
        if v[3] == "Не вышел" or not v[6]: continue
        org_counts[v[6]] += 1
        org_cats[v[6]].add(v[3])

    sorted_orgs = sorted(org_counts.items(), key=lambda x: x[1], reverse=True)
    top_n    = sorted_orgs[:10]
    bottom_n = sorted_orgs[-10:] if len(sorted_orgs) > 10 else []

    ws.merge_cells("A1:D1")
    ws["A1"] = "Топ-10 организаций по визитам"
    ws["A1"].font = _title_font(); ws["A1"].alignment = CENTER; ws["A1"].fill = FILL_GREEN

    for cell in ws[2]:
        _apply_header(cell, FILL_SUBHEAD)
    ws.append(["№", "Организация", "Визитов", "Категории"])
    for cell in ws[2]:
        _apply_header(cell, FILL_SUBHEAD)

    for i, (org, cnt) in enumerate(top_n, 1):
        ws.append([i, org, cnt, ", ".join(org_cats[org])])
        for j, cell in enumerate(ws[i + 2], 1):
            _apply_body(cell, FILL_GREEN if i <= 3 else _fill("FFFFFF"), LEFT if j == 2 else CENTER)

    # Аутсайдеры
    gap_row = ws.max_row + 2
    ws.merge_cells(f"A{gap_row}:D{gap_row}")
    ws[f"A{gap_row}"] = "Аутсайдеры — организации с минимальным числом визитов"
    ws[f"A{gap_row}"].font = _title_font()
    ws[f"A{gap_row}"].fill = FILL_ORANGE
    ws[f"A{gap_row}"].alignment = CENTER

    hdr_row = gap_row + 1
    ws.append(["№", "Организация", "Визитов", "Категории"])
    for cell in ws[hdr_row]:
        _apply_header(cell, FILL_SUBHEAD)

    for i, (org, cnt) in enumerate(reversed(bottom_n), 1):
        ws.append([i, org, cnt, ", ".join(org_cats[org])])
        for j, cell in enumerate(ws[ws.max_row], 1):
            _apply_body(cell, FILL_RED if cnt == 1 else FILL_ORANGE, LEFT if j == 2 else CENTER)

    _set_col_widths(ws, [4, 35, 10, 30])

# ─── ЛИСТ 4: ПО РЕГИОНАМ ──────────────────────────────────────────────────────

def _sheet_regions(wb, all_visits, users):
    ws = wb.create_sheet("📍 Регионы")
    user_map = {u[0]: u[2] for u in users}

    region_data = defaultdict(lambda: defaultdict(int))
    for v in all_visits:
        if v[3] == "Не вышел": continue
        region = user_map.get(v[1], v[2] or "—")
        region_data[region][v[3]] += 1

    ws.merge_cells("A1:F1")
    ws["A1"] = "Распределение визитов по регионам"
    ws["A1"].font = _title_font(); ws["A1"].alignment = CENTER; ws["A1"].fill = FILL_BLUE

    headers = ["Регион", "🩺 Врачи", "💊 Аптеки", "🚚 Оптом", "Итого", "% от всех"]
    ws.append(headers)
    for cell in ws[2]:
        _apply_header(cell)

    grand_total = sum(sum(cats.values()) for cats in region_data.values())

    for idx, (region, cats) in enumerate(sorted(region_data.items()), 3):
        doc = cats.get("🩺 Врач", 0)
        ph  = cats.get("💊 Аптека", 0)
        di  = cats.get("🚚 Дистрибьютор", 0)
        tot = doc + ph + di
        pct = f"{_pct(tot, grand_total)}%"

        ws.append([region, doc, ph, di, tot, pct])
        fill = FILL_ALT if idx % 2 == 0 else _fill("FFFFFF")
        for cell in ws[idx]:
            _apply_body(cell, fill)

    _set_col_widths(ws, [20, 10, 10, 10, 10, 10])
    ws.freeze_panes = "A3"

    # Круговая диаграмма по категориям
    cat_totals = defaultdict(int)
    for v in all_visits:
        if v[3] != "Не вышел": cat_totals[v[3]] += 1

    pie_start = ws.max_row + 2
    ws[f"A{pie_start}"] = "Категория"
    ws[f"B{pie_start}"] = "Кол-во"
    ws[f"A{pie_start}"].font = Font(name="Arial", bold=True)
    ws[f"B{pie_start}"].font = Font(name="Arial", bold=True)

    for i, (cat, cnt) in enumerate(cat_totals.items(), pie_start + 1):
        ws[f"A{i}"] = cat
        ws[f"B{i}"] = cnt

    pie = PieChart()
    pie.title  = "Распределение по категориям"
    pie.style  = 10
    pie.width  = 16
    pie.height = 14

    data = Reference(ws, min_col=2, min_row=pie_start, max_row=pie_start + len(cat_totals))
    cats_ref = Reference(ws, min_col=1, min_row=pie_start + 1, max_row=pie_start + len(cat_totals))
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats_ref)
    ws.add_chart(pie, "H2")

# ─── ЛИСТ 5: ДИНАМИКА ПО ДНЯМ ─────────────────────────────────────────────────

def _sheet_trend(wb, all_visits, start_dt, end_dt):
    ws = wb.create_sheet("📈 Динамика")

    ws.merge_cells("A1:C1")
    ws["A1"] = "Динамика визитов по дням"
    ws["A1"].font = _title_font(); ws["A1"].alignment = CENTER; ws["A1"].fill = FILL_BLUE

    headers = ["Дата", "День недели", "Визитов"]
    ws.append(headers)
    for cell in ws[2]:
        _apply_header(cell)

    day_counts = defaultdict(int)
    for v in all_visits:
        if v[3] == "Не вышел": continue
        dt = _parse_dt(v[4])
        if dt: day_counts[dt.strftime("%d.%m.%Y")] += 1

    current = start_dt
    row_idx = 3
    while current <= end_dt:
        date_str = current.strftime("%d.%m.%Y")
        day_name = DAYS_RU[current.weekday()]
        count    = day_counts.get(date_str, 0)
        ws.append([date_str, day_name, count])

        fill = FILL_ALT if row_idx % 2 == 0 else _fill("FFFFFF")
        if current.weekday() >= 5: fill = _fill("F0F0F0")
        for cell in ws[row_idx]:
            _apply_body(cell, fill)

        current  += timedelta(days=1)
        row_idx  += 1

    _set_col_widths(ws, [14, 12, 12])
    ws.freeze_panes = "A3"

    # Линейный график
    line = LineChart()
    line.title         = "Динамика визитов по дням"
    line.y_axis.title  = "Кол-во визитов"
    line.x_axis.title  = "Дата"
    line.style         = 10
    line.width         = 24
    line.height        = 14

    data = Reference(ws, min_col=3, min_row=2, max_row=row_idx - 1)
    cats = Reference(ws, min_col=1, min_row=3, max_row=row_idx - 1)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    ws.add_chart(line, "E2")

# ─── ЛИСТ 6: ТОП И АУТСАЙДЕРЫ ВРАЧИ ─────────────────────────────────────────

def _sheet_doctors(wb, all_visits):
    ws = wb.create_sheet("👨‍⚕️ Врачи")

    doctor_counts = defaultdict(lambda: {"visits": 0, "orgs": set()})
    for v in all_visits:
        if v[3] != "🩺 Врач" or not v[5]: continue
        doctor_counts[v[5]]["visits"] += 1
        if v[6]: doctor_counts[v[5]]["orgs"].add(v[6])

    sorted_docs = sorted(doctor_counts.items(), key=lambda x: x[1]["visits"], reverse=True)
    top_docs    = sorted_docs[:15]
    bottom_docs = sorted_docs[-10:] if len(sorted_docs) > 15 else []

    ws.merge_cells("A1:D1")
    ws["A1"] = "Топ-15 врачей по числу визитов"
    ws["A1"].font = _title_font(); ws["A1"].alignment = CENTER; ws["A1"].fill = FILL_GREEN

    ws.append(["№", "ФИО врача", "Визитов", "Организация(и)"])
    for cell in ws[2]:
        _apply_header(cell, FILL_SUBHEAD)

    for i, (name, data) in enumerate(top_docs, 1):
        ws.append([i, name, data["visits"], ", ".join(data["orgs"])])
        fill = FILL_GREEN if i <= 3 else (_fill("FFFFFF") if i % 2 else FILL_ALT)
        for j, cell in enumerate(ws[i + 2], 1):
            _apply_body(cell, fill, LEFT if j in (2, 4) else CENTER)

    gap = ws.max_row + 2
    ws.merge_cells(f"A{gap}:D{gap}")
    ws[f"A{gap}"] = "Аутсайдеры — врачи с минимальным числом визитов"
    ws[f"A{gap}"].font = _title_font()
    ws[f"A{gap}"].fill = FILL_ORANGE
    ws[f"A{gap}"].alignment = CENTER

    hdr = gap + 1
    ws.append(["№", "ФИО врача", "Визитов", "Организация(и)"])
    for cell in ws[hdr]:
        _apply_header(cell, FILL_SUBHEAD)

    for i, (name, data) in enumerate(reversed(bottom_docs), 1):
        ws.append([i, name, data["visits"], ", ".join(data["orgs"])])
        for j, cell in enumerate(ws[ws.max_row], 1):
            _apply_body(cell, FILL_RED if data["visits"] == 1 else FILL_ORANGE,
                        LEFT if j in (2, 4) else CENTER)

    _set_col_widths(ws, [4, 30, 10, 40])

# ─── ГЛАВНАЯ ФУНКЦИЯ ГЕНЕРАЦИИ ────────────────────────────────────────────────

def generate_analytics_excel(visits_db, users_db, start_dt, end_dt, period_name, filename):
    users      = fetch_all_users(users_db)
    all_visits = fetch_visits_range(visits_db, start_dt, end_dt)

    visits_by_user = defaultdict(list)
    for v in all_visits:
        visits_by_user[v[1]].append(v)

    if start_dt.tzinfo:
        start_naive = start_dt.replace(tzinfo=None)
        end_naive   = end_dt.replace(tzinfo=None)
    else:
        start_naive, end_naive = start_dt, end_dt

    working_days = _working_days(start_naive, end_naive)
    if period_name == "день":    working_days = 1
    elif period_name == "неделю": working_days = 5

    wb = openpyxl.Workbook()

    _sheet_main(wb, users, visits_by_user, working_days, period_name, start_naive, end_naive)
    _sheet_weekdays(wb, visits_by_user, users)
    _sheet_orgs(wb, all_visits)
    _sheet_regions(wb, all_visits, users)
    _sheet_trend(wb, all_visits, start_naive, end_naive)
    _sheet_doctors(wb, all_visits)

    wb.save(filename)
    return filename
