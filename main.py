import os
import re
import json
import requests
import openpyxl
from io import BytesIO
from urllib.parse import quote
from datetime import datetime

# ================= КОНФИГУРАЦИЯ =================
REPO_OWNER = "colderuopen-art"
REPO_NAME = "raspisanie"
FILE_PATH = "Расписание.xlsx"
TARGET_GROUP = "183р"
STATE_FILE = "state.json"

TG_BOT_TOKEN = os.getenv("TG_BOT_TOKEN")
TG_CHAT_ID = os.getenv("TG_CHAT_ID")

# Сетка звонков по расписанию
BELL_SCHEDULE = {
    "ПОНЕДЕЛЬНИК": {
        "0 пара": "08:30 – 09:20",
        "1 пара": "09:20 – 10:40",
        "2 пара": "10:50 – 12:50",
        "3 пара": "13:00 – 14:20",
        "4 пара": "14:30 – 15:50",
    },
    "DEFAULT": {  # Вторник – Пятница
        "1 пара": "08:30 – 09:50",
        "2 пара": "10:00 – 12:00",
        "3 пара": "12:10 – 13:30",
        "4 пара": "13:40 – 15:00",
        "5 пара": "15:10 – 16:30",
        "6 пара": "16:40 – 18:00",
    },
    "СУББОТА": {
        "1 пара": "08:30 – 09:50",
        "2 пара": "10:10 – 11:30",
        "3 пара": "11:50 – 13:10",
        "4 пара": "13:20 – 14:40",
        "5 пара": "14:50 – 16:10",
        "6 пара": "16:20 – 17:40",
    }
}

# Обеды для литеры "Р"
LUNCH_SCHEDULE = {
    "ПОНЕДЕЛЬНИК": "🥪 <i>Обед: 11:30 – 12:10</i>",
    "DEFAULT": "🥪 <i>Обед: 11:20 – 12:00</i>"
}

NUM_EMOJI = {
    "0": "0️⃣", "1": "1️⃣", "2": "2️⃣",
    "3": "3️⃣", "4": "4️⃣", "5": "5️⃣", "6": "6️⃣"
}

VALID_DAYS = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА"]


def get_latest_commit_sha():
    """Проверка последнего коммита для файла Excel через GitHub API"""
    url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/commits?path={quote(FILE_PATH)}&page=1&per_page=1"
    res = requests.get(url, timeout=15)
    if res.status_code == 200:
        data = res.json()
        if data:
            return data[0]["sha"]
    return None


def download_excel():
    """Скачивание файла в оперативную память"""
    url = f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/main/{quote(FILE_PATH)}"
    res = requests.get(url, timeout=30)
    if res.status_code != 200:
        # Резервная ветка master
        url = f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/master/{quote(FILE_PATH)}"
        res = requests.get(url, timeout=30)
    res.raise_for_status()
    return BytesIO(res.content)


def parse_schedule(file_bytes):
    """Извлечение расписания для группы 183Р по всем вкладкам"""
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    schedule_data = {}

    for sheet_name in wb.sheetnames:
        clean_day = sheet_name.strip().upper()
        if clean_day not in VALID_DAYS:
            continue

        sheet = wb[sheet_name]

        # 1. Поиск даты в первой строке
        date_str = ""
        for cell in sheet[1]:
            val = str(cell.value or "").strip()
            date_match = re.search(r"\d{1,2}\.\d{1,2}\.\d{2,4}", val)
            if date_match:
                date_str = date_match.group(0)
                break

        # 2. Поиск колонок с номерами пар во второй строке
        pair_columns = {}
        for col_idx, cell in enumerate(sheet[2], start=1):
            val = str(cell.value or "").strip().lower()
            if "пара" in val:
                pair_columns[col_idx] = cell.value.strip()

        # 3. Поиск строки с группой 183р в столбце A
        group_row_idx = None
        for row_idx in range(3, sheet.max_row + 1):
            cell_val = str(sheet.cell(row=row_idx, column=1).value or "").strip().lower()
            if cell_val == TARGET_GROUP or cell_val.startswith(TARGET_GROUP):
                group_row_idx = row_idx
                break

        if not group_row_idx:
            continue

        # 4. Сбор только тех пар, которые фактически стоят (без пустых)
        day_pairs = {}
        for col_idx, pair_name in pair_columns.items():
            cell_val = sheet.cell(row=group_row_idx, column=col_idx).value
            if cell_val:
                clean_lesson = str(cell_val).strip()
                if clean_lesson and clean_lesson.lower() != "none":
                    day_pairs[pair_name] = clean_lesson

        schedule_data[clean_day] = {
            "date": date_str,
            "pairs": day_pairs
        }

    return schedule_data


def format_telegram_message(day_name, day_info, changes):
    """Генерация красивого HTML-сообщения"""
    date_str = day_info.get("date", "")
    header_date = f"{day_name}, {date_str}".strip(", ")

    # Блок изменений
    diff_lines = []
    for pair_name, (old_val, new_val) in changes.items():
        if old_val and new_val:
            diff_lines.append(f"• <b>{pair_name}:</b> <s>{old_val}</s> ➔ <b>{new_val}</b>")
        elif not old_val and new_val:
            diff_lines.append(f"• <b>{pair_name}:</b> <i>добавлена</i> ➔ <b>{new_val}</b>")
        elif old_val and not new_val:
            diff_lines.append(f"• <b>{pair_name}:</b> <s>{old_val}</s> ➔ <i>отменена</i>")

    diff_text = "\n".join(diff_lines)

    # Выбор звонков
    if day_name == "ПОНЕДЕЛЬНИК":
        bells = BELL_SCHEDULE["ПОНЕДЕЛЬНИК"]
        lunch = LUNCH_SCHEDULE["ПОНЕДЕЛЬНИК"]
    elif day_name == "СУББОТА":
        bells = BELL_SCHEDULE["СУББОТА"]
        lunch = None
    else:
        bells = BELL_SCHEDULE["DEFAULT"]
        lunch = LUNCH_SCHEDULE["DEFAULT"]

    # Блок пар (без «пар нет», только существующие)
    schedule_lines = []
    for pair_name, lesson in sorted(day_info.get("pairs", {}).items()):
        pair_num = pair_name.split()[0]
        emoji = NUM_EMOJI.get(pair_num, "🔹")
        time_range = bells.get(pair_name, "")
        time_part = f"<b>{time_range}</b> | " if time_range else ""
        
        schedule_lines.append(f"{emoji} {time_part}{lesson}")
        
        # Добавляем обед после 2-й пары
        if pair_name == "2 пара" and lunch:
            schedule_lines.append(f"    ↳ {lunch}")

    schedule_text = "\n".join(schedule_lines)
    now_time = datetime.now().strftime("%d.%m.%Y в %H:%M")

    return (
        f"🔔 <b>Внимание! Изменение в расписании</b>\n\n"
        f"👥 <b>Группа:</b> {TARGET_GROUP.upper()}\n"
        f"📅 <b>{header_date}</b>\n\n"
        f"⚠️ <b>ЧТО ИЗМЕНИЛОСЬ:</b>\n"
        f"{diff_text}\n\n"
        f"───────────────────\n"
        f"📋 <b>АКТУАЛЬНОЕ РАСПИСАНИЕ НА ДЕНЬ:</b>\n"
        f"{schedule_text}\n\n"
        f"<i>(Обновлено: {now_time})</i>"
    )


def send_telegram_notification(text):
    """Отправка сообщения в Telegram"""
    if not TG_BOT_TOKEN or not TG_CHAT_ID:
        print("Ошибка: Токен или ID чата не настроены!")
        return
    url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TG_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True
    }
    res = requests.post(url, json=payload, timeout=10)
    res.raise_for_status()


def main():
    # 1. Загрузка старого состояния
    state = {}
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
        except Exception:
            state = {}

    last_sha = state.get("last_commit_sha")
    old_schedule = state.get("schedule", {})

    # 2. Быстрая проверка изменений файла в репозитории
    current_sha = get_latest_commit_sha()
    if current_sha and current_sha == last_sha:
        print("Файл не обновлялся в репозитории. Выход.")
        return

    # 3. Скачиваем и парсим
    print("Обнаружен новый коммит или первый запуск! Скачиваем Excel...")
    excel_bytes = download_excel()
    new_schedule = parse_schedule(excel_bytes)

    # 4. Первый запуск — просто сохраняем данные без спама
    if not old_schedule:
        print("Первый запуск: сохраняем первичное состояние.")
        state["last_commit_sha"] = current_sha
        state["schedule"] = new_schedule
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        return

    # 5. Поиск точечных изменений именно для 183р
    has_changes = False
    for day, new_info in new_schedule.items():
        old_info = old_schedule.get(day, {"pairs": {}})
        old_pairs = old_info.get("pairs", {})
        new_pairs = new_info.get("pairs", {})

        all_pair_names = set(old_pairs.keys()).union(new_pairs.keys())
        day_changes = {}

        for p_name in all_pair_names:
            old_val = old_pairs.get(p_name)
            new_val = new_pairs.get(p_name)
            if old_val != new_val:
                day_changes[p_name] = (old_val, new_val)

        if day_changes:
            has_changes = True
            print(f"Найдены изменения для 183р на день: {day}")
            msg = format_telegram_message(day, new_info, day_changes)
            send_telegram_notification(msg)

    if not has_changes:
        print("Файл изменился, но пары группы 183Р не затронуты. Сообщение не отправлено.")

    # 6. Обновляем локальное состояние
    state["last_commit_sha"] = current_sha
    state["schedule"] = new_schedule
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
