"""Универсальное извлечение кандидатов событий из строевого приказа.

Пример:
  python extract_order.py "2026\\2026-09-27 №280 (Черкашин).docx"
"""

from __future__ import annotations

import argparse
import csv
import re
from datetime import date, timedelta
from pathlib import Path

from docx import Document
from openpyxl import load_workbook


ROOT = Path(r"C:\WORK\Windows\Строевые приказы")
SHPO_PATH = ROOT / "ШПО.xlsx"
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


def load_shpo() -> tuple[dict[str, str], dict[str, str]]:
    book = load_workbook(SHPO_PATH, read_only=True, data_only=True)
    alf = {text_id(row[0]): str(row[1]) for row in book["АЛФ"].iter_rows(min_row=2, values_only=True) if row[0] and row[1]}
    sheet = book["ОС"]
    headers = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
    ipn_column, rank_column = headers.index("ІПН"), headers.index("Військове звання фактично")
    ranks = {text_id(row[ipn_column]): str(row[rank_column]) for row in sheet.iter_rows(min_row=2, values_only=True) if row[ipn_column] and row[rank_column]}
    return alf, ranks


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
    alf, shpo_ranks = load_shpo()
    rows: list[dict[str, str]] = []
    context = ""
    section = ""
    inherited_date = order_day.strftime("%d.%m.%Y")

    for index, text in enumerate(paragraphs):
        # Заголовки разделов сохраняются до следующего заголовка и задают контекст для персоналий.
        text = text.replace("\u00a0", " ")
        heading = re.match(r"^(\d+(?:\.\d+)*)\.\s", text)
        if heading:
            section = heading.group(1)
            context = text
            inherited_date = order_day.strftime("%d.%m.%Y")
        if not text.lower().startswith("підстава:") and not IPN_RE.search(text):
            inherited_date = event_date(text, datetime_date(inherited_date))
        # Персоналия начинается со звания; назначения ТВО не являются движением.
        person_text = re.sub(r"^\d+(?:\.\d+)*\.\s*", "", text)
        ipn_match = IPN_RE.search(text)
        rank_match = re.match(r"(солдата запасу|рядового запасу|майстер-сержанта|головного сержанта|молодшого лейтенанта|старшого солдата|молодшого сержанта|старшого сержанта|старшого лейтенанта|сержанта|лейтенанта|майора|солдата|солдат|матроса)\s", person_text, re.I)
        if not rank_match:
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
        full_text = " ".join([text, *continuation])
        if "зарахувати до списків особового складу" in full_text.lower():
            category = ("Список", "Зарахувати до списку")
        elif "виключений зі списків особового складу" in text.lower():
            category = ("Список", "Виключити зі списку")
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
            category = ("Відсутні", "Самовільне залишення частини")
        elif "у відрядження:" in context.lower():
            category = ("Відсутні", "Вибув у відрядження")
        elif "вибули зі складу сил" in context.lower() and "." not in section:
            category = ("Прибули", "Вибув до постійного місця служби")
        elif "госпіталізований" in text.lower():
            category = ("Відсутні", "Госпіталізований з відпустки для лікування")
        elif "нижчепойменованих військовослужбовців зарахувати на продовольче забезпечення" in context.lower():
            category = ("Прибули", "Зарахувати на продовольче забезпечення")
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
        event_source = full_text
        if category[1] == "Виключити зі списку":
            event_source = re.split(r"вважати таким, що", text, flags=re.I)[-1]
        row["Дата події"] = event_date(event_source, datetime_date(inherited_date))
        explicit_food = re.search(r"продовольчого забезпечення з\s+([^.]*)", text, re.I)
        row["Продовольча дата"] = event_date("з " + explicit_food.group(1), order_day) if explicit_food else ""
        if category[1] == "Самовільне залишення частини":
            row["Продовольча дата"] = row["Дата події"]
        if not ipn:
            manual = {"МОМОТА": ("МОМОТ Денис Васильович", "майстер-сержант"), "ОСТИМЧУКА": ("ОСТИМЧУК Артем Вікторович", "головний сержант"), "ФІЛІПОВА": ("ФІЛІПОВ Андрій Олександрович", "головний сержант"), "КОВТЮХА": ("КОВТЮХ Олександр Валерійович", "солдат"), "КУЗНЕЦОВА": ("КУЗНЕЦОВ Олександр Юрійович", "сержант"), "ТЕСЛОВА": ("ТЕСЛОВ Максим Миколайович", "старший солдат"), "ІКАВЦЯ": ("ІКАВЕЦЬ Володимир Тарасович", "старший солдат")}
            for surname, values in manual.items():
                if surname in text:
                    row["ПІБ"], row["Звання"] = values
        if "ТРУФАНОВА" in text and not ipn:
            row["ПІБ"], row["Звання"] = "ТРУФАНОВ Сергій Олександрович", "майор"
        row["Звання"] = shpo_ranks.get(ipn, row["Звання"] or rank_match.group(1).replace("лейтенанта", "лейтенант").replace("молодшого", "молодший").replace("майстер-сержанта", "майстер-сержант"))
        if category == ("Список", "Зарахувати до списку") and "тимчасово прибулого особового складу" in text.lower():
            close = row.copy()
            close["Лист"], close["Дія"], close["Деталі"] = "Прибули", "Закрити ТП", "Зараховано до списків особового складу"
            rows.append(close)

    fields = ["Лист", "Дія", "Звання", "ПІБ", "ІПН", "Дата події", "Наказ", "Деталі", "Текст наказу", "Підстава", "Супровідний документ", "Термін", "Дорога", "Продовольча дата"]
    with output_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Приказ №{number}: {len(rows)} событий -> {output_path}")
    for sheet in ("Список", "Відсутні", "Прибули"):
        print(f"{sheet}: {sum(row['Лист'] == sheet for row in rows)}")


if __name__ == "__main__":
    main()