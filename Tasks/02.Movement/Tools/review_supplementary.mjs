// Отдельные журналы ТВО и питания получают проверенные строки, без выдуманных переходов.
export const schemas = {
  "ТВО": ["Звання", "ПІБ", "ІПН", "УІД", "Посада.Своя", "Посада.Заміщує", "Термін", "Доступ.Наказ", "Прийняття.Наказ", "Прийняття", "Здавання.План", "Здавання", "Здавання.Наказ", "Прийняття.Підстава", "Здавання.Підстава", "Додаткова інформація", "Текст наказу"],
  "Продовольче": ["Звання", "ПІБ", "Установа", "Група", "Зарахування", "Зарахування.Наказ", "Виключення", "Виключення.Наказ", "Зарахування.Підстава", "Виключення.Підстава", "Додаткова інформація", "Текст наказу"],
  "Сухпрод": ["Звання", "ПІБ", "Установа", "Група", "Початок", "Тривалість", "Припинення", "Підстава", "Наказ", "Додаткова інформація", "Текст наказу"],
  "Проблеми": ["Тип", "ПІБ / ідентифікатор", "Лист / пункт / рядок", "Проблема", "Значення в джерелах", "Що перевірити"]
};
const dateHeaders = new Set(["Прийняття", "Здавання.План", "Здавання", "Зарахування", "Виключення", "Початок", "Припинення"]);
for (const name of ["ТВО", "Продовольче", "Сухпрод"]) schemas[name].unshift("Дія");
const movementDates = new Set([...dateHeaders, "Вибуття.ПЗ", "Вибуття", "Прибуття.План", "Прибуття", "Прибуття.ПЗ", "Вибуття.План"]);
const empty = value => value === undefined || value === null || value === "";
const serial = value => {
  if (typeof value !== "string" || !/^\d{2}\.\d{2}\.\d{4}$/.test(value)) return value;
  const [day, month, year] = value.split(".").map(Number);
  return (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000;
};

export function prepareSupplementary(payload, records, orderNumber) {
  if (payload.notes?.length && !Array.isArray(payload.problems)) {
    throw new Error("Старые notes нельзя переносить целиком: подготовьте problems только с конкретными расхождениями; обычный журнал оставьте в Work");
  }
  if (payload.food?.length) {
    throw new Error("Старый массив food требует ручной сверки ПІБ и званий: передайте проверенные объекты dryRations для Сухпрод");
  }
  if (records.some(row => !empty(row["ТВО.ПІБ"])) && !Array.isArray(payload.tvo)) {
    throw new Error("Подготовьте отдельный массив tvo по всем упоминаниям ТВО в приказе, включая сдачу и возвращение к своим обязанностям");
  }
  const groups = {
    "ТВО": [...(payload.tvo ?? [])],
    "Продовольче": [...(payload.foodSupport ?? [])],
    "Сухпрод": [...(payload.dryRations ?? [])],
    "Проблеми": [...(payload.problems ?? [])]
  };
  for (const row of records.filter(row => !empty(row["ТВО.ПІБ"]))) {
    if (!groups["ТВО"].some(tvo => (row["ТВО.ІПН"] && tvo["ІПН"] === row["ТВО.ІПН"]) || tvo["ПІБ"] === row["ТВО.ПІБ"])) throw new Error(`ТВО не перенесено в отдельный журнал: ${row["ТВО.ПІБ"]}`);
  }
  for (let index = records.length - 1; index >= 0; index--) {
    const row = records[index];
    if (groups[row["Лист"]]) {
      groups[row["Лист"]].unshift(row);
      records.splice(index, 1);
      continue;
    }
    // Только отдельные операции питания уходят из журнала физических прибытий.
    if (row["Лист"] === "Прибули" && /продовольч/i.test(row["Дія"] ?? "") && !/Закрити ТП/.test(row["Дія"])) {
      if (!empty(row["Дата події"])) throw new Error(`Уточните физическую дату отдельно от питания: ${row["ПІБ"]}`);
      const closing = /^(Вибув|Зняти|Виключити)/.test(row["Дія"]);
      const date = row["Продовольча дата"];
      if (!date || date === "НЕ ЗМІНЮВАТИ") throw new Error(`Нет явной даты операции питания: ${row["ПІБ"]}`);
      const prefix = closing ? "Виключення" : "Зарахування";
      groups["Продовольче"].unshift({ "Звання": row["Звання"], "ПІБ": row["ПІБ"], "Установа": row["Звідки"], [prefix]: date, [`${prefix}.Наказ`]: Number(orderNumber), [`${prefix}.Підстава`]: row["Підстава"], "Текст наказу": row["Текст наказу"] });
      records.splice(index, 1);
    }
  }
  for (const [name, rows] of Object.entries(groups)) {
    for (const row of rows) {
      if (!row || Array.isArray(row) || typeof row !== "object") throw new Error(`${name}: ожидаются объекты с именами колонок`);
      if (name === "Проблеми" && (!row["Проблема"] || !row["Що перевірити"])) throw new Error("Проблеми: укажите конкретное расхождение и действие для сверки");
      if (name === "ТВО" && (!row["Дія"] || !row["ПІБ"] || !row["Текст наказу"])) throw new Error("ТВО: нужны проверенные действие, полный ПІБ и пункт приказа");
      if (name === "ТВО" && !empty(row["УІД"])) throw new Error("УІД результата должна оставаться пустой");
      if (name === "Продовольче" && (empty(row["Зарахування"]) === empty(row["Виключення"]))) throw new Error("Продовольче: одна строка должна содержать одну операцию зачисления или снятия");
      if (name === "Сухпрод") {
        if (!row["ПІБ"] || empty(row["Початок"]) || !Number.isInteger(Number(row["Тривалість"])) || Number(row["Тривалість"]) <= 0) throw new Error("Сухпрод: нужны ПІБ, начало и положительное целое число суток");
        row["Тривалість"] = Number(row["Тривалість"]);
      }
      if (["Продовольче", "Сухпрод"].includes(name) && !empty(row["Додаткова інформація"])) throw new Error("В журналах питания Додаткова інформація остаётся пустой; проблемы передавайте отдельно");
    }
  }
  return groups;
}

export function applyLayout(sheet, headers, values) {
  const widthFor = header => {
    if (header === "ПІБ") return 20;
    if (header === "Звання") return 10;
    if (header === "Дія") return 14;
    if (header === "ІПН") return 12;
    if (header === "ІН") return 20;
    if (movementDates.has(header)) return 12;
    if (/Посада/.test(header)) return 12;
    if (header === "Подія") return 17;
    if (["Куди", "Звідки"].includes(header)) return 20;
    if (/Підстава/.test(header)) return 28;
    if (/інформація/.test(header)) return 24;
    if (header === "Супровідний документ") return 14;
    if (header === "Текст наказу") return 120;
    if (header === "Установа") return 12;
    if (header === "Проблема" || header === "Що перевірити" || header === "Значення в джерелах") return 38;
    if (header === "ПІБ / ідентифікатор") return 24;
    if (header === "Лист / пункт / рядок") return 22;
    if (header === "Тип") return 20;
    return 9.25;
  };
  const widths = headers.map(header => ["УІД", "За межі"].includes(header) ? 5 : widthFor(header));
  // Основные поля помещаем в ширину экрана; полный исходный пункт остаётся справа.
  const core = headers.map((header, index) => header === "Текст наказу" ? 0 : widths[index]);
  const total = core.reduce((sum, width) => sum + width, 0);
  if (total > 290) {
    const adjustable = core.map((width, index) => ({ width, index })).filter(item => item.width > 14 && headers[item.index] !== "Додаткова інформація");
    const available = adjustable.reduce((sum, item) => sum + item.width - 14, 0);
    for (const item of adjustable) widths[item.index] -= Math.min(total - 290, available) * (item.width - 14) / available;
  }
  headers.forEach((header, index) => {
    sheet.getRangeByIndexes(0, index, Math.max(values.length + 4, 5), 1).format.columnWidth = widths[index];
    if (["ІПН", "ІН", "УІД", "Посада", "Посада.Своя", "Посада.Заміщує"].includes(header)) sheet.getRangeByIndexes(4, index, Math.max(values.length, 1), 1).setNumberFormat("@");
  });
  const range = sheet.getRangeByIndexes(3, 0, Math.max(values.length + 1, 2), headers.length);
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  range.format.wrapText = true;
  headers.forEach((header, index) => {
    if (header === "ІПН" || movementDates.has(header)) sheet.getRangeByIndexes(4, index, Math.max(values.length, 1), 1).format.wrapText = false;
  });
  const headerLines = Math.max(...headers.map((header, index) => Math.ceil(header.length / Math.max(2, widths[index] * 1.1))));
  sheet.getRangeByIndexes(3, 0, 1, headers.length).format.rowHeight = Math.max(42, Math.min(180, headerLines * 11 + 10));
}

export function addSupplementary(workbook, groups) {
  const ranges = {};
  const tableNames = { "ТВО": "ActingDutyEvents", "Продовольче": "FoodSupportEvents", "Сухпрод": "DryRationEvents", "Проблеми": "ReviewProblems" };
  for (const name of ["ТВО", "Продовольче", "Сухпрод", "Проблеми"]) {
    const sheet = workbook.worksheets.add(name);
    sheet.showGridLines = false;
    const headers = schemas[name];
    const rows = groups[name];
    const values = rows.map(row => headers.map(header => {
      if (header === "Дія") return row[header] ?? (name === "Продовольче" ? (row["Виключення"] ? "Виключити з продовольчого забезпечення" : "Зарахувати на продовольче забезпечення") : name === "Сухпрод" ? "Видати сухпайок" : "");
      const value = row[header] ?? "";
      if (dateHeaders.has(header)) return serial(value);
      if (/Наказ$/.test(header) && /^\d+$/.test(String(value))) return Number(value);
      return value;
    }));
    const lastCol = String.fromCharCode(64 + headers.length);
    const bottom = Math.max(5, values.length + 4);
    sheet.getRange(`A1:${lastCol}${bottom + 4}`).format = { fill: "#262626", font: { name: "Times New Roman", size: 10, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center" };
    sheet.getRange(`A1:${lastCol}1`).merge();
    sheet.getRange("A1").values = [[name]];
    sheet.getRange(`A1:${lastCol}1`).format = { fill: "#404040", font: { name: "Times New Roman", size: 14, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", rowHeight: 25 };
    sheet.getRange(`A4:${lastCol}4`).values = [headers];
    if (values.length) sheet.getRangeByIndexes(4, 0, values.length, headers.length).values = values;
    const table = sheet.tables.add(`A4:${lastCol}${bottom}`, true, tableNames[name]);
    table.style = "TableStyleDark1";
    table.showFilterButton = true;
    sheet.getRange(`A4:${lastCol}4`).format = { fill: "#000000", font: { name: "Times New Roman", size: 10, bold: true, color: "#FFFFFF" }, rowHeight: 42 };
    for (let index = 0; index < values.length; index++) {
      sheet.getRangeByIndexes(index + 4, 0, 1, headers.length).format.fill = index % 2 ? "#383838" : "#262626";
    }
    headers.forEach((header, index) => {
      if (dateHeaders.has(header)) sheet.getRangeByIndexes(4, index, Math.max(values.length, 1), 1).setNumberFormat("dd.mm.yyyy");
    });
    if (name === "Сухпрод" && values.length) {
      const stop = headers.indexOf("Припинення");
      sheet.getRangeByIndexes(4, stop, values.length, 1).formulas = values.map(() => ["=[@Початок]+[@Тривалість]"]);
      // Расчётное прекращение является заполненной колонкой, даже если вход содержит только начало и срок.
      values.forEach(row => { row[stop] = Number(row[headers.indexOf("Початок")]) + Number(row[headers.indexOf("Тривалість")]); });
    }
    applyLayout(sheet, headers, values);
    sheet.getRange("A1:AZ1000").format.fill = "#28251E";
    sheet.getRange(`A1:${lastCol}1`).format.fill = "#5D5142";
    sheet.getRange(`A4:${lastCol}4`).format.fill = "#5D5142";
    for (let index = 0; index < values.length; index++) sheet.getRangeByIndexes(index + 4, 0, 1, headers.length).format.fill = index % 2 ? "#38342C" : "#302D25";
    for (let index = 0; index < values.length; index++) {
      const lines = Math.max(...values[index].map((value, column) => empty(value) ? 1 : Math.ceil(String(value).length / Math.max(3, sheet.getRangeByIndexes(index + 4, column, 1, 1).format.columnWidth))));
      sheet.getRangeByIndexes(index + 4, 0, 1, headers.length).format.rowHeight = Math.min(409, Math.max(36, lines * 12 + 8));
    }
    ranges[name] = `A1:${lastCol}${Math.min(bottom, 10)}`;
  }
  return ranges;
}