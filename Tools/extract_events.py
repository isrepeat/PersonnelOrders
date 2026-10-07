from docx import Document
from pathlib import Path
import csv
import re

REPOSITORY_ROOT = Path(__file__).resolve().parent.parent
ROOT = REPOSITORY_ROOT / "Documents" / "Orders" / "А7383" / "2026"
OUTPUT = REPOSITORY_ROOT / "events_212_277.csv"
MONTHS = {"січня": "01", "лютого": "02", "березня": "03", "квітня": "04", "травня": "05", "червня": "06", "липня": "07", "серпня": "08", "вересня": "09", "жовтня": "10", "листопада": "11", "грудня": "12"}
ORDER_RE = re.compile(r"№\s*(\d+)")
NAME_RE = re.compile(r"([А-ЯІЇЄҐ][А-ЯІЇЄҐ'’\-]{2,}\s+[А-ЯІЇЄҐ][а-яіїєґ'’\-]+\s+[А-ЯІЇЄҐ][а-яіїєґ'’\-]+(?:\s+[А-ЯІЇЄҐ][а-яіїєґ'’\-]+)?)")
IPN_RE = re.compile(r"\b(\d{10})\b")
DATE_RE = re.compile(r"(?:з\s+)?(\d{1,2})\s+(%s)\s+(2026)\s+року" % "|".join(MONTHS))
RANGE_START_RE = re.compile(r"з\s+(\d{1,2})\s+(%s)\s+по\s+\d{1,2}\s+(?:%s)\s+(2026)\s+року" % ("|".join(MONTHS), "|".join(MONTHS)))


def iso_date(text):
    range_match = RANGE_START_RE.search(text.lower())
    if range_match:
        return f"{range_match.group(3)}-{MONTHS[range_match.group(2)]}-{int(range_match.group(1)):02d}"
    match = DATE_RE.search(text.lower())
    if not match:
        return ""
    return f"{match.group(3)}-{MONTHS[match.group(2)]}-{int(match.group(1)):02d}"


def section_kind(text):
    lower = text.lower()
    if "тимчасово прибул" in lower and "зарахувати" in lower:
        return "Прибули", "Зарахувати до ТП"
    if "зарахувати до списків" in lower:
        return "Список", "Зарахувати до списку"
    if "виключити із списків" in lower or "виключити зі списків" in lower or "виключити з списків" in lower:
        return "Список", "Виключити зі списку"
    if "виключити з тимчасово прибул" in lower or "вибув з тимчасово прибул" in lower:
        return "Прибули", "Вибув з ТП"
    if "вважати таким, що вибув" in lower or "вважати таким, який вибув" in lower or "не повернувся" in lower or "самовільно залишив" in lower or "зниклим безвісти" in lower or "зняти з усіх видів забезпечення" in lower:
        return "Відсутні", "Вибув"
    return None, None


def return_kind(heading, text):
    lower = (heading + " " + text).lower()
    if "з лікувального закладу" in lower:
        return "Відсутні", "Прибув з лікування"
    if "з щорічної основної відпустки" in lower:
        return "Відсутні", "Прибув з щорічної відпустки"
    if "з відпустки за сімейними" in lower:
        return "Відсутні", "Прибув з відпустки за сімейними обставинами"
    if "з відпустки для лікування" in lower:
        return "Відсутні", "Прибув з відпустки для лікування"
    if "із самовільного залишення" in lower:
        return "Відсутні", "Повернувся із СЗЧ"
    if "з відрядження" in lower:
        return "Відсутні", "Прибув з відрядження"
    return None, None


rows = []
for path in sorted(ROOT.glob("*.docx")):
    if path.name.startswith("~$"):
        continue
    order_match = ORDER_RE.search(path.name)
    if not order_match or int(order_match.group(1)) < 212:
        continue
    order_number = int(order_match.group(1))
    order_date = path.name[:10]
    document = Document(path)
    paragraphs = [paragraph.text.strip().replace("\t", " ") for paragraph in document.paragraphs]
    heading = ""
    for index, paragraph in enumerate(paragraphs):
        if not paragraph or paragraph.startswith("Підстава:") or paragraph.startswith("Станом на"):
            continue
        if re.match(r"^\d+(?:\.\d+)*\.?\s", paragraph):
            heading = paragraph
        sheet, action = section_kind(paragraph)
        if not sheet:
            sheet, action = return_kind(heading, paragraph)
        if not sheet:
            continue
        if paragraph.startswith("Поновити") or paragraph.startswith("Призупинити"):
            continue
        names = NAME_RE.findall(paragraph)
        for name in names:
            if name.split()[0] in {"МІНІСТЕРСТВО", "НАЦІОНАЛЬНОЇ", "ЗБРОЙНИХ", "ХАРКІВСЬКОЇ", "КОМАНДИР"}:
                continue
            ipn_match = IPN_RE.search(paragraph)
            event_date = iso_date(paragraph) or iso_date(heading) or order_date
            rows.append({
                "Лист": sheet,
                "Дія": action,
                "Дата події": event_date,
                "ПІБ як у наказі": name,
                "РНОКПП": ipn_match.group(1) if ipn_match else "",
                "Наказ": f"№{order_number} від {order_date}",
                "Пункт": heading[:280],
                "Текст для перевірки": paragraph[:950],
                "Файл": path.name,
            })

seen = set()
unique_rows = []
for row in rows:
    key = tuple(row.values())
    if key not in seen:
        seen.add(key)
        unique_rows.append(row)

fields = list(unique_rows[0]) if unique_rows else ["Лист", "Дія", "Дата події", "ПІБ як у наказі", "РНОКПП", "Наказ", "Пункт", "Текст для перевірки", "Файл"]
with OUTPUT.open("w", newline="", encoding="utf-8-sig") as file:
    writer = csv.DictWriter(file, fieldnames=fields)
    writer.writeheader()
    writer.writerows(unique_rows)
print(f"Подій: {len(unique_rows)}")
for sheet in ("Список", "Відсутні", "Прибули"):
    print(f"{sheet}: {sum(row['Лист'] == sheet for row in unique_rows)}")