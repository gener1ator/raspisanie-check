import os
import re
import json
import time
import hashlib
import requests
import openpyxl
from io import BytesIO
from urllib.parse import quote
from datetime import datetime, date, time as dtime, timedelta, timezone

# ================= КОНФИГУРАЦИЯ =================
REPO_OWNER = "colderuopen-art"
REPO_NAME = "raspisanie"
FILE_PATH = "Расписание.xlsx"
TARGET_GROUP = "183р"
STATE_FILE = "state.json"

TG_BOT_TOKEN = os.getenv("TG_BOT_TOKEN")
TG_CHAT_ID = os.getenv("TG_CHAT_ID")

# Часовой пояс (по умолчанию UTC+3, Москва). Можно изменить в Secrets репозитория.
TIMEZONE_OFFSET = int(os.getenv("TIMEZONE_OFFSET", "5"))

# Сетка звонков
BELL_SCHEDULE = {
    "ПОНЕДЕЛЬНИК": {
        "0 пара": "08:30 – 09:20",
        "1 пара": "09:20 – 10:40",
        "2 пара": "10:50 – 12:50",
        "3 пара": "13:00 – 14:20",
        "4 пара": "14:30 – 15:50",
        "5 пара": "16:00 – 17:20",
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

# Обеды для группы 183Р (литера "Р")
LUNCH_SCHEDULE = {
    "ПОНЕДЕЛЬНИК": "🥪 <i>Обед: 11:30 – 12:10</i>",
    "DEFAULT": "🥪 <i>Обед: 11:20 – 12:00</i>"
}

NUM_EMOJI = {
    "0": "0️⃣", "1": "1️⃣", "2": "2️⃣",
    "3": "3️⃣", "4": "4️⃣", "5": "5️⃣", "6": "6️⃣"
}

VALID_DAYS = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА"]
DAY_NAMES_RU = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ"]


def get_local_now():
    """Текущее локальное время с учетом часового пояса"""
    tz = timezone(timedelta(hours=TIMEZONE_OFFSET))
    return datetime.now(tz)


def get_latest_commit_sha():
    """Проверка последнего коммита через GitHub API"""
    url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/commits?path={quote(FILE_PATH)}&page=1&per_page=1"
    try:
        res = requests.get(url, timeout=15)
        if res.status_code == 200:
            data = res.json()
            if data:
                return data[0]["sha"]
    except Exception as e:
        print(f"Ошибка проверки коммита: {e}")
    return None


def download_excel():
    """Скачивание Excel-файла"""
    url = f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/main/{quote(FILE_PATH)}"
    res = requests.get(url, timeout=30)
    if res.status_code != 200:
        url = f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/master/{quote(FILE_PATH)}"
        res = requests.get(url, timeout=30)
    res.raise_for_status()
    return BytesIO(res.content)


def extract_date_from_sheet(sheet):
    """Извлечение даты из первых строк вкладки"""
    for row in sheet.iter_rows(min_row=1, max_row=3, values_only=False):
        for cell in row:
            val = cell.value
            if not val:
                continue
            if isinstance(val, (datetime, date)):
                return val.strftime("%d.%m.%Y")
            val_str = str(val).strip()
            match_dot = re.search(r"\b\d{1,2}\.\d{1,2}\.\d{2,4}\b", val_str)
            if match_dot:
                return match_dot.group(0)
            match_iso = re.search(r"\b\d{4}-\d{1,2}-\d{1,2}\b", val_str)
            if match_iso:
                try:
                    dt = datetime.strptime(match_iso.group(0), "%Y-%m-%d")
                    return dt.strftime("%d.%m.%Y")
                except Exception:
                    return match_iso.group(0)
    return ""


def parse_schedule(file_bytes):
    """Сбор расписания для группы 183Р"""
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    schedule_data = {}

    for sheet_name in wb.sheetnames:
        clean_day = sheet_name.strip().upper()
        if clean_day not in VALID_DAYS:
            continue

        sheet = wb[sheet_name]
        date_str = extract_date_from_sheet(sheet)

        pair_columns = {}
        for col_idx, cell in enumerate(sheet[2], start=1):
            val = str(cell.value or "").strip().lower()
            if "пара" in val:
                pair_columns[col_idx] = cell.value.strip()

        group_row_idx = None
        for row_idx in range(3, sheet.max_row + 1):
            cell_val = str(sheet.cell(row=row_idx, column=1).value or "").strip().lower()
            if cell_val == TARGET_GROUP or cell_val.startswith(TARGET_GROUP):
                group_row_idx = row_idx
                break

        if not group_row_idx:
            continue

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


def get_pair_end_time(day_name, pair_name):
    """Получение времени окончания пары в виде объекта datetime.time"""
    if day_name == "ПОНЕДЕЛЬНИК":
        bells = BELL_SCHEDULE["ПОНЕДЕЛЬНИК"]
    elif day_name == "СУББОТА":
        bells = BELL_SCHEDULE["СУББОТА"]
    else:
        bells = BELL_SCHEDULE["DEFAULT"]

    time_range = bells.get(pair_name, "")
    if time_range:
        parts = re.split(r"[–-]", time_range)
        if len(parts) == 2:
            end_str = parts[1].strip()
            try:
                h, m = map(int, end_str.split(":"))
                return dtime(h, m)
            except Exception:
                pass
    return dtime(15, 0)


def determine_target_day(schedule, now):
    """
    Определяет, на какой день сейчас должен висеть закреп:
    - если сегодня учебный день и пары еще идут -> СЕГОДНЯ
    - если сегодня пары кончились (или выходной) -> СЛЕДУЮЩИЙ УЧЕБНЫЙ ДЕНЬ
    """
    weekday_idx = now.weekday()
    today_name = DAY_NAMES_RU[weekday_idx]
    current_time = now.time()

    today_info = schedule.get(today_name, {})
    today_pairs = today_info.get("pairs", {})

    # Если сегодня учебный день и у нас есть пары
    if today_name in VALID_DAYS and today_pairs:
        # Находим время конца последней пары за сегодня
        end_times = [get_pair_end_time(today_name, p) for p in today_pairs.keys()]
        latest_end_time = max(end_times) if end_times else dtime(15, 0)

        # Если последняя пара еще не закончилась -> закрепляем СЕГОДНЯ
        if current_time < latest_end_time:
            return today_name, "СЕГОДНЯ"

    # Иначе ищем ближайший следующий учебный день
    for offset in range(1, 7):
        next_idx = (weekday_idx + offset) % 7
        candidate_day = DAY_NAMES_RU[next_idx]
        candidate_info = schedule.get(candidate_day, {})
        if candidate_day in VALID_DAYS and candidate_info.get("pairs"):
            label = "ЗАВТРА" if offset == 1 else candidate_day
            return candidate_day, label

    return "ПОНЕДЕЛЬНИК", "ПОНЕДЕЛЬНИК"


def build_schedule_block(day_name, day_info):
    """Список пар со временем (только существующие пары)"""
    if day_name == "ПОНЕДЕЛЬНИК":
        bells = BELL_SCHEDULE["ПОНЕДЕЛЬНИК"]
        lunch = LUNCH_SCHEDULE["ПОНЕДЕЛЬНИК"]
    elif day_name == "СУББОТА":
        bells = BELL_SCHEDULE["СУББОТА"]
        lunch = None
    else:
        bells = BELL_SCHEDULE["DEFAULT"]
        lunch = LUNCH_SCHEDULE["DEFAULT"]

    pairs = day_info.get("pairs", {})
    if not pairs:
        return "<i>На этот день пар нет</i>"

    def sort_key(item):
        num_part = item[0].split()[0]
        return int(num_part) if num_part.isdigit() else 99

    sorted_pairs = sorted(pairs.items(), key=sort_key)

    lines = []
    for pair_name, lesson in sorted_pairs:
        pair_num = pair_name.split()[0]
        emoji = NUM_EMOJI.get(pair_num, "🔹")
        time_range = bells.get(pair_name, "")
        time_part = f"<b>{time_range}</b> | " if time_range else ""

        lines.append(f"{emoji} {time_part}{lesson}")
        if pair_name == "2 пара" and lunch:
            lines.append(f"    ↳ {lunch}")

    return "\n".join(lines)


def format_pinned_message(day_name, day_info, label_text, now):
    """Текст для закрепленного сообщения"""
    date_str = day_info.get("date", "")
    header_date = f"{day_name}, {date_str}".strip(", ")
    schedule_text = build_schedule_block(day_name, day_info)
    time_str = now.strftime("%H:%M")

    return (
        f"📌 <b>РАСПИСАНИЕ НА {label_text.upper()}</b>\n\n"
        f"👥 <b>Группа:</b> {TARGET_GROUP.upper()}\n"
        f"📆 <b>{header_date}</b>\n\n"
        f"───────────────────\n"
        f"📋 <b>АКТУАЛЬНЫЕ ПАРЫ:</b>\n"
        f"{schedule_text}\n\n"
        f"<i>(Закреплено автоматически • {time_str})</i>"
    )


def format_replacement_message(day_name, day_info, changes, now):
    """Текст уведомления о замене пар"""
    date_str = day_info.get("date", "")
    header_date = f"{day_name}, {date_str}".strip(", ")

    diff_lines = []
    for pair_name, (old_val, new_val) in sorted(changes.items()):
        if old_val and new_val:
            diff_lines.append(f"• <b>{pair_name}:</b> <s>{old_val}</s> ➔ <b>{new_val}</b>")
        elif not old_val and new_val:
            diff_lines.append(f"• <b>{pair_name}:</b> <i>добавлена</i> ➔ <b>{new_val}</b>")
        elif old_val and not new_val:
            diff_lines.append(f"• <b>{pair_name}:</b> <s>{old_val}</s> ➔ <i>отменена</i>")

    diff_text = "\n".join(diff_lines)
    schedule_text = build_schedule_block(day_name, day_info)
    now_time = now.strftime("%d.%m.%Y в %H:%M")

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


def send_telegram_message(text):
    """Отправка сообщения в Telegram с возвратом ID сообщения"""
    if not TG_BOT_TOKEN or not TG_CHAT_ID:
        return None
    url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TG_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True
    }
    try:
        res = requests.post(url, json=payload, timeout=10)
        res.raise_for_status()
        return res.json().get("result", {}).get("message_id")
    except Exception as e:
        print(f"Ошибка отправки сообщения: {e}")
        return None


def edit_telegram_message(message_id, text):
    """Бесшумное редактирование сообщения"""
    if not TG_BOT_TOKEN or not TG_CHAT_ID or not message_id:
        return False
    url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/editMessageText"
    payload = {
        "chat_id": TG_CHAT_ID,
        "message_id": message_id,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True
    }
    try:
        res = requests.post(url, json=payload, timeout=10)
        res.raise_for_status()
        return True
    except Exception as e:
        print(f"Ошибка редактирования сообщения {message_id}: {e}")
        return False


def pin_telegram_message(message_id):
    """Закрепление сообщения в чате"""
    if not TG_BOT_TOKEN or not TG_CHAT_ID or not message_id:
        return
    url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/pinChatMessage"
    payload = {
        "chat_id": TG_CHAT_ID,
        "message_id": message_id,
        "disable_notification": True  # Закреплять тихо, без громкого звука
    }
    try:
        res = requests.post(url, json=payload, timeout=10)
        res.raise_for_status()
        print(f"Сообщение {message_id} успешно закреплено!")
    except Exception as e:
        print(f"Ошибка закрепления: проверьте права бота на закреп в чате! ({e})")


def main():
    now = get_local_now()

    # 1. Загружаем память
    state = {}
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
        except Exception:
            state = {}

    last_sha = state.get("last_commit_sha")
    old_schedule = state.get("schedule", {})

    # 2. Проверяем коммит в репозитории
    current_sha = get_latest_commit_sha()
    excel_downloaded = False
    new_schedule = old_schedule

    if current_sha and current_sha != last_sha:
        print("Обнаружен новый коммит. Скачиваем Excel...")
        try:
            excel_bytes = download_excel()
            new_schedule = parse_schedule(excel_bytes)
            excel_downloaded = True
            state["last_commit_sha"] = current_sha
        except Exception as e:
            print(f"Ошибка парсинга Excel: {e}")

    # 3. Отправка уведомлений о заменах (если файл обновлялся)
    if excel_downloaded and old_schedule:
        for day, new_info in new_schedule.items():
            old_info = old_schedule.get(day, {"pairs": {}, "date": ""})
            old_pairs = old_info.get("pairs", {})
            new_pairs = new_info.get("pairs", {})

            old_date = old_info.get("date", "").strip()
            new_date = new_info.get("date", "").strip()

            # Точечные замены
            all_pair_names = set(old_pairs.keys()).union(new_pairs.keys())
            day_changes = {}
            for p_name in all_pair_names:
                old_val = old_pairs.get(p_name)
                new_val = new_pairs.get(p_name)
                if old_val != new_val:
                    day_changes[p_name] = (old_val, new_val)

            if day_changes and new_date == old_date:
                print(f"Замена на {day} для 183Р!")
                msg = format_replacement_message(day, new_info, day_changes, now)
                send_telegram_message(msg)
                time.sleep(1)

    # 4. ЛОГИКА ЗАКРЕПА (выполняется при каждом запуске!)
    if new_schedule:
        target_day, label_text = determine_target_day(new_schedule, now)
        target_info = new_schedule.get(target_day)

        if target_info:
            pinned_text = format_pinned_message(target_day, target_info, label_text, now)
            text_hash = hashlib.md5(pinned_text.encode("utf-8")).hexdigest()

            saved_pinned_day = state.get("pinned_day")
            saved_msg_id = state.get("pinned_message_id")
            saved_hash = state.get("pinned_text_hash")

            # Случай А: День закрепа сменился (учебный день закончился) или закрепа еще нет
            if saved_pinned_day != target_day or not saved_msg_id:
                print(f"Переключение закрепа на: {target_day} ({label_text})")
                new_msg_id = send_telegram_message(pinned_text)
                if new_msg_id:
                    pin_telegram_message(new_msg_id)
                    state["pinned_day"] = target_day
                    state["pinned_message_id"] = new_msg_id
                    state["pinned_text_hash"] = text_hash

            # Случай Б: День тот же, но поменялись пары -> просто редактируем закреп
            elif saved_hash != text_hash:
                print(f"Обновление текста в существующем закрепе {saved_msg_id}")
                success = edit_telegram_message(saved_msg_id, pinned_text)
                if success:
                    state["pinned_text_hash"] = text_hash
                else:
                    # Если сообщение удалили руками, шлем и крепим заново
                    new_msg_id = send_telegram_message(pinned_text)
                    if new_msg_id:
                        pin_telegram_message(new_msg_id)
                        state["pinned_message_id"] = new_msg_id
                        state["pinned_text_hash"] = text_hash

    # 5. Сохраняем состояние
    state["schedule"] = new_schedule
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
