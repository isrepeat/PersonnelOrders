"""Универсальное извлечение кандидатов событий из строевого приказа.

Пример:
  python Tools/extract_order.py "2026\\2026-09-27 №280 (Черкашин).docx"
"""

from __future__ import annotations

import argparse
import csv
import re
from datetime import date, timedelta
from pathlib import Path

from docx import Document
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent.parent
SHPO_PATH = ROOT / "Tools" / "ШПО.xlsx"
IPN_RE = re.compile(r"\b(\d{10})\b")
RANK_RE = re.compile(r"(?:^|\.\s*)(солдата запасу|рядового|солдата|солдат|матроса|старшого солдата|молодшого сержанта|старшого сержанта|сержанта|старшого лейтенанта|лейтенанта|майора)\s+", re.I)
NAME_RE = re.compile(r"(?:солдата запасу|рядового|солдата|солдат|матроса|старшого солдата|молодшого сержанта|старшого сержанта|сержанта|старшого лейтенанта|лейтенанта|майора)\s+([А-ЯІЇЄҐ'’-]+)\s+([А-ЯІЇЄҐ][а-яіїєґ'’-]+)\s+([А-ЯІЇЄҐ][а-яіїєґ'’-]+)", re.I)
MONTHS = {"січня": 1, "лютого": 2, "березня": 3, "квітня": 4, "травня": 5, "червня": 6, "липня": 7, "серпня": 8, "вересня": 9, "жовтня": 10, "листопада": 11, "грудня": 12}
RANKS = {"солдата запасу": "солдат", "рядового": "рядовий", "солдата": "солдат", "солдат": "солдат", "матроса": "матрос", "старшого солдата": "старший солдат", "молодшого сержанта": "молодший сержант", "старшого сержанта": "старший сержант", "сержанта": "сержант", "старшого лейтенанта": "старший лейтенант", "лейтенанта": "лейтенант", "майора": "майор"}


def text_id(value: object) -> str:
    """Возвращает ИПН без преобразования в научную нотацию или десятичную дробь."""
    if value is None:
        return ""
    return str(value).split(".")[0]


def position_label(value: object) -> str:
    """Заменяет специальные коды посады условными обозначениями учёта."""
    code = str(value) if value is not None else ""
    for prefix, label in (("A1A", "РОЗП"), ("A1B", "СПИС"), ("A1C", "ТП")):
        if code.startswith(prefix):
            return label
    return code


def load_shpo() -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    book = load_workbook(SHPO_PATH, read_only=True, data_only=True)
    alf = {text_id(row[0]): str(row[1]) for row in book["АЛФ"].iter_rows(min_row=2, values_only=True) if row[0] and row[1]}
    sheet = book["ОС"]
    headers = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    ipn_column, rank_column = headers.index("ІПН"), headers.index("Військове звання фактично")
    ranks = {text_id(row[ipn_column]): str(row[rank_column]) for row in sheet.iter_rows(min_row=2, values_only=True) if row[ipn_column] and row[rank_column]}
    position_column = headers.index("Код посади")
    positions = {text_id(row[ipn_column]): position_label(row[position_column]) for row in sheet.iter_rows(min_row=2, values_only=True) if row[ipn_column] and row[position_column]}
    book.close()
    return alf, ranks, positions


def acting_replacement(text: str, alf: dict[str, str], positions: dict[str, str]) -> tuple[dict[str, str], str]:
    """Извлекает назначенного ТВО; сопоставляет все три компонента ПІБ, не одну фамилию."""
    result = {key: "" for key in ("ТВО.ПІБ", "ТВО.ІПН", "ТВО.Посада")}
    assignments = list(re.finditer(
        r"(?i:тимчасов[ео]\s+виконання\s+обов[’']?язків)[^.!?]*?"
        r"(?i:покласти\s+на)\s+[^.;]{0,80}?"
        r"([А-ЯІЇЄҐ][А-ЯІЇЄҐ’'-]+)\s+([А-ЯІЇЄҐ][а-яіїєґ’'-]+)\s+([А-ЯІЇЄҐ][а-яіїєґ’'-]+)",
        text.replace("\u00a0", " "),
    ))
    if not assignments:
        return result, ""
    if len(assignments) != 1:
        return result, "ТВО: несколько назначений в одном пункте, требуется сверка"
    assignment = assignments[0]
    original = assignment.groups()

    def forms(word: str, index: int) -> set[str]:
        word = word.casefold().replace("’", "'")
        variants = {word, word + "а", word + "у", word + "я", word + "ю"}
        if word.endswith("й"):
            variants.update({word[:-1] + "я", word[:-1] + "ю"})
        if index == 0 and word.endswith(("ий", "ій")):
            variants.update({word[:-2] + "ого", word[:-2] + "ього"})
        if word.endswith("о"):
            variants.update({word[:-1] + "а", word[:-1] + "у"})
        if word.endswith("а"):
            variants.update({word[:-1] + "и", word[:-1] + "і"})
        if word.endswith("ь"):
            variants.update({word[:-1] + "я", word[:-1] + "ю"})
        if index == 1:
            irregular = {"павло": "павла", "лев": "лева", "ігор": "ігоря"}
            if word in irregular:
                variants.add(irregular[word])
        return variants

    explicit = re.match(r"\s*,?\s*(?:ІПН\s*[:–-]?\s*)?(\d{10})\b", text[assignment.end():], re.I)
    if explicit:
        ipn = explicit.group(1)
        matches = [ipn] if ipn in alf else []
    else:
        matches = []
        for ipn, name in alf.items():
            parts = name.split()
            if len(parts) == 3 and all(
                actual.casefold().replace("’", "'") in forms(expected, index)
                for index, (actual, expected) in enumerate(zip(original, parts))
            ):
                matches.append(ipn)
    if len(matches) != 1:
        return result, f"ТВО: не найдено однозначное соответствие ШПО для {' '.join(original)}"
    ipn = matches[0]
    result.update({"ТВО.ПІБ": alf[ipn], "ТВО.ІПН": ipn, "ТВО.Посада": position_label(positions.get(ipn, ""))})
    return result, "" if result["ТВО.Посада"] else f"ТВО: в ШПО отсутствует код посады для {alf[ipn]}"


def nominative_word(word: str, index: int) -> str:
    lower = word.lower()
    if index == 0 and lower.endswith("енка"):
        return word[:-1] + "О"
    if index == 0 and lower.endswith(("ова", "ева", "іна", "їна")):
        return word[:-1]
    if index == 1 and lower.endswith(("а", "я", "у", "ю")):
        return word[:-1] + ("й" if lower.endswith(("я", "ю")) else "")
    if index == 2 and lower.endswith(("овича", "йовича", "євича", "ича")):
        return word[:-1]
    return word


def fallback_name(text: str) -> str:
    text = re.sub(r"(?:рекрута|головного сержанта|штаб-сержанта|майстер-сержанта|капітана(?: медичної служби)?)\s+", "солдата ", text, flags=re.I)
    match = NAME_RE.search(text)
    return " ".join(nominative_word(word, index) for index, word in enumerate(match.groups())) if match else ""


def parse_order_info(paragraphs: list[str]) -> tuple[str, date]:
    source = " ".join(paragraphs[:15])
    match = re.search(r"(\d{1,2})\.(\d{1,2})\.(20\d{2}).{0,40}?№\s*(\d+)", source)
    if not match:
        raise ValueError("Не найдены дата и номер приказа в первых абзацах.")
    day, month, year, number = map(int, match.groups())
    return str(number), date(year, month, day)


def datetime_date(value: str) -> date:
    day, month, year = map(int, value.split("."))
    return date(year, month, day)


def event_date(text: str, default: date) -> str:
    text = text.replace("\u00a0", " ")
    interval = re.search(r"з\s+(\d{1,2})\s+по\s+\d{1,2}\s+([а-яіїєґ]+)\s+(20\d{2})", text, re.I)
    if interval and interval.group(2).lower() in MONTHS:
        day, month_word, year = interval.groups()
        return date(int(year), MONTHS[month_word.lower()], int(day)).strftime("%d.%m.%Y")
    match = re.search(r"з\s+(\d{1,2})\s+([а-яіїєґ]+)(?:\s+(20\d{2}))?", text, re.I)
    if not match or match.group(2).lower() not in MONTHS:
        return default.strftime("%d.%m.%Y")
    day, month_word, year = match.groups()
    return date(int(year or default.year), MONTHS[month_word.lower()], int(day)).strftime("%d.%m.%Y")


def basis_after(paragraphs: list[str], start: int) -> str:
    for text in paragraphs[start + 1:start + 8]:
        if text.startswith("Підстава:"):
            return text
    return ""


def clean_basis(text: str) -> tuple[str, str]:
    documents = ", ".join(re.findall(r"\b20\d{2}/\d+/(?:\d+(?:-20\d{2}/\d+/\d+)?)\b", text))
    result = text.removeprefix("Підстава:").strip()
    result = re.sub(r"[;,]?\s*(?:відпускний квиток|посвідчення про відрядження).*?(?=;|$)", "", result, flags=re.I)
    return re.sub(r"\s+", " ", result).strip(" ;."), documents


def departure_details(text: str) -> tuple[str, str]:
    term = re.search(r"терміном на\s+([^,.]+)", text, re.I)
    road = re.search(r"(\d+\s+діб[^,.]*проїзд[^,.]*)", text, re.I)
    return term.group(1) if term else "", road.group(1) if road else ""


def destination(text: str) -> str:
    """Отделяет направление движения от срока, даты и цели поездки."""
    text = re.sub(r"\s+", " ", text.replace("\u00a0", " ")).strip()
    personal = re.search(r"(?:направити у відрядження|госпіталізований|вибув у відпустку для лікування)\s+(?:у|в)\s+(.+)", text, re.I)
    if not personal:
        personal = re.search(r",\s+(?:у|в)\s+(.+)", text, re.I)
    if not personal:
        personal = re.search(r",\s+((?:м|с|смт)\.\s+.+)", text, re.I)
    if personal:
        text = personal.group(1)
    elif re.match(r"^(?:у|в|з|із|до)\s+(?!\d|відрядження|відпустки|лікувального закладу)", text, re.I):
        text = re.sub(r"^(?:у|в|з|із|до)\s+", "", text, flags=re.I)
    else:
        return ""
    text = re.split(r",?\s*(?:терміном|без урахування|з метою|до окремого розпорядження|з\s+\d{1,2}\s+[а-яіїєґ]+\s+20\d{2})|\.\s+(?:Видати|Поновити|Направити)", text, maxsplit=1, flags=re.I)[0].strip(" ,:.")
    unit = re.match(r"військов(?:у|ої)\s+частин[уи]\s+([АA]?\d{4})\b", text, re.I)
    return unit.group(1) if unit else text


def absence_type(action: str) -> str:
    """Возвращает тип отсутствия, а не служебное действие открытия/закрытия."""
    mapping = {
        "Прибув з відрядження": "Відрядження", "Вибув у відрядження": "Відрядження",
        "Прибув з лікування": "Стаціонарне лікування", "Вибув на лікування": "Стаціонарне лікування",
        "Госпіталізований з відпустки для лікування": "Стаціонарне лікування",
        "Прибув зі щорічної відпустки": "Щорічна відпустка", "Вибув у щорічну відпустку": "Щорічна відпустка",
        "Прибув з відпустки за сімейними обставинами": "Відпустка за сімейними обставинами",
        "Вибув у відпустку за сімейними обставинами": "Відпустка за сімейними обставинами",
        "Вибув у відпустку для лікування": "Відпустка для лікування",
        "Прибув із СЗЧ": "Самовільне залишення частини", "Закрити СЗЧ": "Самовільне залишення частини",
        "Самовільне залишення частини": "Самовільне залишення частини",
        "Не повернувся з відпустки": "Самовільне залишення частини", "Не повернувся з лікування": "Самовільне залишення частини",
        "Безвісти зниклий": "Безвісти зниклий", "Повернення до іншої частини": "Повернення до іншої частини",
    }
    return mapping.get(action, "")


def return_before_departure(row: dict[str, str]) -> list[dict[str, str]]:
    """Закрывает СЗЧ фактическим возвратом и выделяет промежуток до командировки."""
    if row["Лист"] != "Відсутні" or row["Дія"] != "Вибув у відрядження":
        return []
    text = row["Текст наказу"]
    if not re.search(r"самовільн\w*\s+залиш", text, re.I):
        return []
    date_pattern = r"(\d{1,2}\s+[а-яіїєґ]+\s+20\d{2})\s+року"
    returned = re.search(date_pattern + r"\s+повернувся\s+до\s+військової\s+частини\s+([АA]\d{4})", text, re.I)
    sent = re.search(r"направити\s+(?:у|в)\s+відрядження\s+(?:у|в)\s+військову\s+частину\s+([АA]\d{4})\s+з\s+" + date_pattern, text, re.I)
    if not returned or not sent:
        return []
    order_day = datetime_date(re.search(r"\d{2}\.\d{2}\.\d{4}", row["Наказ"]).group(0))
    return_day = event_date("з " + returned.group(1), order_day)
    departure_day = event_date("з " + sent.group(2), order_day)
    if datetime_date(return_day) > datetime_date(departure_day):
        row["Деталі"] = "Проверить: дата возвращения после даты направления в командировку"
        return []
    row.update({"Дата події": departure_day, "Куди": sent.group(1).replace("A", "А")})
    closed = row.copy()
    closed.update({"Дія": "Закрити СЗЧ", "Дата події": return_day, "Прибуття": return_day,
                   "Куди": "", "Термін": "", "Дорога": "", "Супровідний документ": ""})
    for key in ("ТВО.ПІБ", "ТВО.ІПН", "ТВО.Посада"):
        closed[key] = ""
    if return_day == departure_day:
        return [closed]
    bridge = closed.copy()
    bridge.update({"Дія": "Повернення до іншої частини", "Куди": returned.group(2).replace("A", "А"),
                   "Вибуття": return_day, "Прибуття": departure_day})
    if row.get("Продовольча дата") != "НЕ ЗМІНЮВАТИ":
        closed["Продовольча дата"] = row.get("Продовольча дата") or max(datetime_date(return_day), order_day + timedelta(days=1)).strftime("%d.%m.%Y")
        bridge["Вибуття.Продовольче"] = closed["Продовольча дата"]
        bridge["Прибуття.Продовольче"] = row.get("Продовольча дата") or max(datetime_date(departure_day), order_day + timedelta(days=1)).strftime("%d.%m.%Y")
    return [closed, bridge]


def exclusion_absences(row: dict[str, str]) -> list[dict[str, str]]:
    """Закрывает СЗЧ и выделяет период после возвращения в другую часть."""
    if row["Лист"] != "Список" or row["Дія"] != "Виключити зі списку":
        return []
    text = row["Текст наказу"]
    returned = re.search(
        r"який\s+(\d{1,2}\s+[а-яіїєґ]+\s+20\d{2})\s+року\s+повернувся\s+до\s+військової\s+частини\s+([АA]\d{4})",
        text, re.I,
    )
    if not returned or "після самовільного залишення" not in text.lower():
        return []
    exclusion_day = datetime_date(row["Дата події"])
    return_day = event_date("з " + returned.group(1), exclusion_day)
    closing_day = (exclusion_day + timedelta(days=1)).strftime("%d.%m.%Y")
    closed = row.copy()
    closed.update({
        "Лист": "Відсутні", "Дія": "Закрити СЗЧ", "Дата події": return_day,
        "Продовольча дата": closing_day, "Прибуття": return_day,
        "Прибуття.Продовольче": closing_day,
        "Деталі": f"Закрити попереднє СЗЧ; повернення до {returned.group(2)}",
    })
    bridge = row.copy()
    bridge.update({
        "Лист": "Відсутні", "Дія": "Повернення до іншої частини",
        "Дата події": return_day, "Продовольча дата": closing_day,
        "Вибуття": return_day, "Вибуття.Продовольче": closing_day,
        "Прибуття": closing_day, "Прибуття.Продовольче": closing_day,
        "Деталі": f"Повернення до {returned.group(2)}; виключення {row['Дата події']}",
        "Куди": returned.group(2),
    })
    return [closed, bridge]


def classify(text: str, context: str) -> tuple[str, str] | None:
    combined = f"{context} {text}".lower()
    if "зарахувати до списків особового складу" in text.lower():
        return "Список", "Зарахувати до списку"
    if "виключити зі списків особового складу" in combined:
        return "Список", "Виключити зі списку"
    if "щорічної основної відпустки" in context.lower() and "вибув" not in combined:
        return "Відсутні", "Прибув зі щорічної відпустки"
    if "відпустки за сімейними" in context.lower() and "вибув" not in combined:
        return "Відсутні", "Прибув з відпустки за сімейними обставинами"
    if "у відрядження" in context.lower() or "у відрядження" in combined:
        return "Відсутні", "Вибув у відрядження"
    if "у частину щорічної" in context.lower() or "вибув у щорічну відпустку" in combined:
        return "Відсутні", "Вибув у щорічну відпустку"
    if "вибув у відпустку для лікування" in combined:
        return "Відсутні", "Вибув у відпустку для лікування"
    if "зняти з котлового забезпечення" in context.lower():
        return "Відсутні", "Зняти з котлового забезпечення"
    return None


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("order", type=Path, help="Путь к .docx приказу, относительно рабочей папки или абсолютный")
    parser.add_argument("--output", type=Path, help="Путь к CSV; по умолчанию events_<номер>.csv")
    args = parser.parse_args()
    order_path = args.order if args.order.is_absolute() else ROOT / args.order
    paragraphs = [paragraph.text.strip().replace("\t", " ") for paragraph in Document(order_path).paragraphs]
    paragraphs = [paragraph for paragraph in paragraphs if paragraph]
    number, order_day = parse_order_info(paragraphs)
    output_path = args.output or ROOT / f"events_{number}.csv"
    alf, shpo_ranks, positions = load_shpo()
    name_ids: dict[str, list[str]] = {}
    for person_ipn, name in alf.items():
        name_ids.setdefault(name, []).append(person_ipn)
    position_names = {name: positions[ids[0]] for name, ids in name_ids.items() if len(ids) == 1 and ids[0] in positions}
    rows: list[dict[str, str]] = []
    context = ""
    section = ""
    inherited_date = order_day.strftime("%d.%m.%Y")
    inherited_destination = ""

    for index, text in enumerate(paragraphs):
        # Заголовки разделов сохраняются до следующего заголовка и задают контекст для персоналий.
        text = text.replace("\u00a0", " ")
        heading = re.match(r"^(\d+(?:\.\d+)*)\.\s", text)
        if heading:
            section = heading.group(1)
            context = text
            inherited_date = order_day.strftime("%d.%m.%Y")
            inherited_destination = ""
        if not text.lower().startswith("підстава:") and not IPN_RE.search(text):
            inherited_date = event_date(text, datetime_date(inherited_date))
        # Персоналия начинается со звания; назначения ТВО не являются движением.
        person_text = re.sub(r"^\d+(?:\.\d+)*\.\s*", "", text)
        ipn_match = IPN_RE.search(text)
        rank_match = re.match(r"(солдата запасу|рядового запасу|рядового|штаб-сержанта|рекрута|капітана медичної служби|капітанa|капітана|майстер-сержанта|головного сержанта|молодшого лейтенанта|старшого солдата|молодшого сержанта|старшого сержанта|старшого лейтенанта|сержанта|лейтенанта|майора|солдата|солдат|матроса)\s", person_text, re.I)
        if not rank_match:
            if not text.lower().startswith("підстава:") and not heading:
                location = destination(text)
                if location:
                    inherited_destination = location
            continue
        # Подраздел задаёт направление события и сбрасывает прежний контекст.
        # Проверяем продолжение пункта: зачисление бывает отдельным абзацем.
        following = []
        for candidate in paragraphs[index + 1:]:
            if candidate.startswith("Підстава:") or re.match(r"^\d+(?:\.\d+)*\.\s", candidate):
                break
            following.append(candidate)
        continuation = []
        for candidate in following:
            if re.match(r"(?:солдата|старшого|молодшого|сержанта|майора|лейтенанта)\s", candidate, re.I):
                break
            if not candidate.lower().startswith("виплачувати"):
                continuation.append(candidate)
        full_text = " ".join([text, *continuation]).replace("\u00a0", " ")
        if "зарахувати до списків особового складу" in full_text.lower():
            category = ("Список", "Зарахувати до списку")
        elif "виключений зі списків особового складу" in full_text.lower() or "виключити зі списків особового складу" in full_text.lower():
            category = ("Список", "Виключити зі списку")
        elif "з відпустки за сімейними" in context.lower():
            category = ("Відсутні", "Прибув з відпустки за сімейними обставинами")
        elif "із самовільного залишення" in context.lower():
            category = ("Відсутні", "Прибув із СЗЧ")
        elif "у відпустку за сімейними" in context.lower():
            category = ("Відсутні", "Вибув у відпустку за сімейними обставинами")
        elif "зниклими безвісти" in context.lower():
            category = ("Відсутні", "Безвісти зниклий")
        elif "не повернувся" in full_text.lower():
            category = ("Відсутні", "Не повернувся з відпустки" if "щорічної відпустки" in text.lower() else "Не повернувся з лікування")
        elif "окрім продовольчого" in full_text.lower() and "направити у відрядження" in full_text.lower():
            category = ("Відсутні", "Вибув у відрядження")
        elif "з лікувального закладу" in context.lower():
            category = ("Відсутні", "Прибув з лікування")
        elif "з відрядження" in context.lower():
            category = ("Відсутні", "Прибув з відрядження")
        elif "з щорічної" in context.lower():
            category = ("Відсутні", "Прибув зі щорічної відпустки")
        elif "у службове відрядження у військову частину а7383" in context.lower():
            category = ("Прибули", "Прибув у відрядження")
        elif "на лікування" in context.lower():
            category = ("Відсутні", "Вибув на лікування")
        elif "у відпустку для лікування" in context.lower():
            category = ("Відсутні", "Вибув у відпустку для лікування")
        elif "у частину щорічної" in context.lower():
            category = ("Відсутні", "Вибув у щорічну відпустку")
        elif "самовільно залишили" in context.lower():
            category = ("Прибули", "Вибув через СЗЧ") if "припинити відрядження" in full_text.lower() else ("Відсутні", "Самовільне залишення частини")
        elif "у відрядження:" in context.lower():
            category = ("Відсутні", "Вибув у відрядження")
        elif "вибули зі складу сил" in context.lower() and "." not in section:
            category = ("Прибули", "Вибув до постійного місця служби")
        elif "госпіталізований" in text.lower():
            category = ("Відсутні", "Госпіталізований з відпустки для лікування")
        elif "нижчепойменованих військовослужбовців зарахувати на продовольче забезпечення" in context.lower():
            category = ("Прибули", "Зарахувати на продовольче забезпечення")
        elif "нижчепойменованих військовослужбовців зняти з продовольчого забезпечення" in context.lower():
            category = ("Прибули", "Вибув з продовольчого забезпечення")
        else:
            category = None
        if not category:
            continue
        ipn = ipn_match.group(1) if ipn_match else ""
        basis_text = ""
        for candidate in paragraphs[index + 1:]:
            if candidate.startswith("Підстава:"):
                basis_text = candidate
                break
            if re.match(r"^\d+(?:\.\d+)*\.\s", candidate):
                break
        basis, companion = clean_basis(basis_text)
        row = {"Лист": category[0], "Дія": category[1], "Звання": shpo_ranks.get(ipn, RANKS.get(rank_match.group(1).lower(), "")), "ПІБ": alf.get(ipn, fallback_name(text)), "ІПН": ipn, "Дата події": event_date(text, order_day), "Наказ": f"№{number} від {order_day.strftime('%d.%m.%Y')}", "Деталі": "", "Текст наказу": text, "Підстава": basis, "Супровідний документ": companion, "Термін": departure_details(text)[0], "Дорога": departure_details(text)[1]}
        rows.append(row)
        row["Текст наказу"] = full_text
        # ТВО назначается при выбытии; возвращение прежнего исполнителя не является новым назначением.
        tvo, tvo_warning = acting_replacement(full_text, alf, positions) if category[0] == "Відсутні" and category[1].startswith("Вибув") else ({key: "" for key in ("ТВО.ПІБ", "ТВО.ІПН", "ТВО.Посада")}, "")
        row.update(tvo)
        if tvo_warning:
            row["Деталі"] = tvo_warning
        row["Куди"] = (destination(full_text) or inherited_destination) if category[0] == "Відсутні" else ""
        event_source = full_text
        if category[1] == "Виключити зі списку":
            exclusion = re.search(r"З\s+\d{1,2}\s+\w+\s+20\d{2}\s+року виключити", full_text, re.I)
            event_source = exclusion.group(0) if exclusion else re.split(r"вважати таким, що", text, flags=re.I)[-1]
        if "окрім продовольчого" in full_text.lower():
            event_source = full_text.split("Направити у відрядження")[-1]
        row["Дата події"] = event_date(event_source, datetime_date(inherited_date))
        explicit_food = re.search(r"продовольчого забезпечення (?:знятий )?з\s+([^.]*)", full_text, re.I)
        row["Продовольча дата"] = event_date("з " + explicit_food.group(1), order_day) if explicit_food else ""
        if category[1] == "Самовільне залишення частини":
            row["Продовольча дата"] = row["Дата події"]
        if "окрім продовольчого" in full_text.lower():
            row["Продовольча дата"] = "НЕ ЗМІНЮВАТИ"
        if not ipn:
            manual = {"МОМОТА": ("МОМОТ Денис Васильович", "майстер-сержант"), "ОСТИМЧУКА": ("ОСТИМЧУК Артем Вікторович", "головний сержант"), "ФІЛІПОВА": ("ФІЛІПОВ Андрій Олександрович", "головний сержант"), "КОВТЮХА": ("КОВТЮХ Олександр Валерійович", "солдат"), "КУЗНЕЦОВА": ("КУЗНЕЦОВ Олександр Юрійович", "сержант"), "ТЕСЛОВА": ("ТЕСЛОВ Максим Миколайович", "старший солдат"), "ІКАВЦЯ": ("ІКАВЕЦЬ Володимир Тарасович", "старший солдат")}
            for surname, values in manual.items():
                if surname in text:
                    row["ПІБ"], row["Звання"] = values
        if "ТРУФАНОВА" in text and not ipn:
            row["ПІБ"], row["Звання"] = "ТРУФАНОВ Сергій Олександрович", "майор"
        row["Звання"] = shpo_ranks.get(ipn, row["Звання"] or rank_match.group(1).replace("лейтенанта", "лейтенант").replace("молодшого", "молодший").replace("майстер-сержанта", "майстер-сержант"))
        # Ручная сверка форм, отсутствующих в справочнике; ключ — полный исходный ПІБ.
        verified_names = {
            "КРАВЦЯ Олега Юрійовича": "КРАВЕЦЬ Олег Юрійович",
            "ПЛУГАТИРЬОВА Павла Вікторовича": "ПЛУГАТИРЬОВ Павло Вікторович",
            "ПОЛТАВЦЯ Володимира Віталійовича": "ПОЛТАВЕЦЬ Володимир Віталійович",
            "ГАРБУЗА Сергія Олександровича": "ГАРБУЗ Сергій Олександрович",
            "ШЕРЕМЕТА Ігоря Ігоровича": "ШЕРЕМЕТ Ігор Ігорович",
            "ПІДГІРНЯКА Романа Андрійовича": "ПІДГІРНЯК Роман Андрійович",
            "СВЄТЛАКОВА Вадима Володимировича": "СВЄТЛАКОВ Вадим Володимирович",
            "ЛЕБЕДЄВА Михайла Миколайовича": "ЛЕБЕДЄВ Михайло Миколайович",
            "МАРКИТАНА Бориса Петровича": "МАРКИТАН Борис Петрович",
        }
        if ipn not in alf:
            for original, nominative in verified_names.items():
                if original in text:
                    row["ПІБ"] = nominative
        if ipn not in shpo_ranks:
            row["Звання"] = {"рекрута": "рекрут", "головного сержанта": "головний сержант", "штаб-сержанта": "штаб-сержант"}.get(rank_match.group(1).lower(), row["Звання"])
        row["Посада"] = positions.get(ipn, "") if ipn else position_names.get(row["ПІБ"], "")
        if category == ("Список", "Зарахувати до списку") and "тимчасово прибулого особового складу" in text.lower():
            close = row.copy()
            close["Лист"], close["Дія"], close["Деталі"] = "Прибули", "Закрити ТП", "Зараховано до списків особового складу"
            rows.append(close)

    # Специальные закрытия добавляются после извлечения исходных событий.
    expanded = []
    for row in rows:
        # Последовательность: закрытие СЗЧ, промежуточное возвращение, новая командировка.
        expanded.extend(return_before_departure(row))
        expanded.append(row)
        expanded.extend(exclusion_absences(row))
    rows = expanded
    for row in rows:
        row["Подія"] = absence_type(row["Дія"]) if row["Лист"] == "Відсутні" else ""
        row["За межі"] = ""
    fields = ["Лист", "Дія", "Звання", "ПІБ", "ІПН", "Дата події", "Наказ", "Деталі", "Текст наказу", "Підстава", "Супровідний документ", "Термін", "Дорога", "Продовольча дата", "Вибуття", "Вибуття.Продовольче", "Прибуття", "Прибуття.Продовольче", "Посада", "Подія", "Куди", "За межі", "ТВО.ПІБ", "ТВО.ІПН", "ТВО.Посада"]
    with output_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Приказ №{number}: {len(rows)} событий -> {output_path}")
    for sheet in ("Список", "Відсутні", "Прибули"):
        print(f"{sheet}: {sum(row['Лист'] == sheet for row in rows)}")


if __name__ == "__main__":
    main()