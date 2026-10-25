import fs from "node:fs/promises";
import { dirname } from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";
import { prepareSupplementary, addSupplementary, applyLayout } from "./review_supplementary.mjs";

// Сборщик РУХ принимает проверенные данные и явный путь новой версии, не перезаписывает старую.
const [inputPath, outputPath, outputDir] = process.argv.slice(2);
if (!inputPath || !outputPath || !outputDir) throw new Error("Укажите входной JSON/CSV, выходной XLSX и папку превью");
try { await fs.access(outputPath); throw new Error(`Файл уже существует: ${outputPath}`); }
catch (error) { if (error.code !== "ENOENT") throw error; }
const payload = inputPath.endsWith('.json') ? JSON.parse(await fs.readFile(inputPath, 'utf8')) : null;
const csv = await fs.readFile(inputPath, "utf8");
const lines = csv.replace(/^\uFEFF/, "").trim().split(/\r?\n/);
const parse = (line) => {
  const cells = [];
  let value = "", quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      if (quoted && line[i + 1] === '"') { value += '"'; i += 1; } else quoted = !quoted;
    } else if (char === "," && !quoted) { cells.push(value); value = ""; } else value += char;
  }
  cells.push(value);
  return cells;
};
const headers = parse(lines[0]);
const records = payload?.records || lines.slice(1).map(parse).map((cells) => Object.fromEntries(headers.map((header, index) => [header, cells[index] ?? ""])));
const orderNumberText = String(payload?.orderNumber ?? records[0]?.["Наказ"]?.match(/№(\d+)/)?.[1] ?? "");
const orderDateText = payload?.orderDate ?? records[0]?.["Наказ"]?.match(/\d{2}\.\d{2}\.\d{4}/)?.[0];
if (!/^\d+$/.test(orderNumberText) || !/^\d{2}\.\d{2}\.\d{4}$/.test(orderDateText ?? "")) throw new Error("Укажите проверенные orderNumber и orderDate");
const [od, om, oy] = orderDateText.split('.').map(Number);
const orderDay = new Date(Date.UTC(oy, om - 1, od));
const nextDay = new Date(orderDay.getTime() + 86400000);
const nextDayText = `${String(nextDay.getUTCDate()).padStart(2, '0')}.${String(nextDay.getUTCMonth()+1).padStart(2, '0')}.${nextDay.getUTCFullYear()}`;
const rankMap = new Map([
  ["солдата", "солдат"],
  ["матроса", "матрос"],
  ["старшого солдата", "старший солдат"],
  ["молодшого сержанта", "молодший сержант"],
  ["старшого сержанта", "старший сержант"],
  ["головного сержанта", "головний сержант"],
  ["старшого лейтенанта", "старший лейтенант"],
]);
const nominativeRank = (record) => {
  const text = record["Текст наказу"] ?? "";
  const sourceRank = [...rankMap.keys()].find((candidate) => new RegExp(`(?:^|\\.\\s*)${candidate}\\s`, "i").test(text));
  return sourceRank ? rankMap.get(sourceRank) : record["Звання"];
};
const supplemental = prepareSupplementary(payload ?? {}, records, orderNumberText);
const workbook = Workbook.create();
const palette = { "Список": "#1F4E78", "Відсутні": "#9E480E", "Прибули": "#548235" };
const tableNames = { "Список": "RosterEvents", "Відсутні": "AbsenceEvents", "Прибули": "ArrivalEvents" };
const order = ["Список", "Відсутні", "Прибули"];

for (const name of order) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  sheet.tabColor = palette[name];
  sheet.getRange("A1:Z1000").format.fill = "#262626";
  const recordsForSheet = records.filter((record) => record["Лист"] === name);
  const absenceHeaders = ["Дія", "Звання", "ПІБ", "ІПН", "Посада", "Подія", "Куди", "За межі", "Вибуття.Наказ", "Вибуття.Продовольче", "Вибуття", "Вибуття.Термін", "Вибуття.Дорога", "Прибуття.План", "Прибуття", "Прибуття.Продовольче", "Прибуття.Наказ", "Вибуття.Підстава", "Супровідний документ", "Прибуття.Підстава", "Текст наказу"];
  const arrivalHeaders = ["Дія", "Звання", "ПІБ", "Посада", "Звідки", "Прибуття.Наказ", "Прибуття.Продовольче", "Прибуття", "Прибуття.Термін", "Вибуття.План", "Вибуття", "Вибуття.Продовольче", "Вибуття.Наказ", "Прибуття.Підстава", "Супровідний документ", "Вибуття.Підстава", "Текст наказу"];
  const rosterHeaders = ["Дія", "Звання", "ПІБ", "ІПН", "Зарахування", "Виключення", "В строю", "На продовольчому", "Зараховано на продовольче", "Виключено з продовольчого", "Підстава", "Супровідний документ", "Текст наказу"];
  let outputHeaders = name === "Відсутні" ? absenceHeaders : name === "Прибули" ? arrivalHeaders : rosterHeaders;
  const uidIndex = outputHeaders.indexOf("ІПН") + 1;
  if (uidIndex) outputHeaders.splice(uidIndex, 0, "УІД");
  if (name === "Прибули") outputHeaders.splice(3, 0, "ІН", "УІД");
  if (name === "Прибули") outputHeaders.splice(outputHeaders.length - 1, 0, "Додаткова інформація");
  const legacyHeaders = [...outputHeaders];
  const selected = name === "Список" ? ["Дія", "ПІБ", "ІПН", "УІД", "Зарахування", "Виключення", "Текст наказу"] : [...outputHeaders];
  if (name === "Відсутні") selected.splice(selected.length - 1, 0, "Додаткова інформація");
  outputHeaders = selected.map(header => header.replace(".Продовольче", ".ПЗ"));
  const lastCol = String.fromCharCode(64 + outputHeaders.length);
  sheet.getRange(`A1:${lastCol}1`).merge();
  sheet.getRange("A1").values = [[`${name}: події для ручного внесення`]];
  sheet.getRange(`A2:${lastCol}2`).merge();
  sheet.getRange("A2").values = [[`Джерело: наказ №${orderNumberText} від ${orderDateText}. ПІБ звірено з ШПО за ІПН; без ІПН приведено до називного відмінка.`]];
  sheet.getRange(`A4:${lastCol}4`).values = [outputHeaders];
  let values = recordsForSheet.map((record) => {
    const rank = record["Звання"];
    const date = record["Дата події"];
    const [day, month, year] = date.split(".").map(Number);
    const eventDate = new Date(Date.UTC(year, month - 1, day));
    const orderDate = orderDay;
    const foodDate = record["Продовольча дата"] === "НЕ ЗМІНЮВАТИ" ? "" : record["Продовольча дата"] || (eventDate > orderDate ? date : nextDayText);
    if (name === "Прибули") {
      const closure = record["Дія"].includes("Закрити ТП") || record["Дія"].startsWith("Вибув");
      return closure
        ? [record["Дія"], rank, record["ПІБ"], record["Посада"] || "", record["Звідки"] || "", "", "", "", "", "", date, foodDate, orderNumberText, "", record["Супровідний документ"], record["Підстава"], record["Текст наказу"]]
        : [record["Дія"], rank, record["ПІБ"], record["Посада"] || "", record["Звідки"] || "", orderNumberText, foodDate, date, "", "", "", "", "", record["Підстава"], record["Супровідний документ"], "", record["Текст наказу"]];
    }
    if (name === "Список") {
      if (record["Дія"].startsWith("Скасувати виключення")) {
        return [record["Дія"], rank, record["ПІБ"], record["ІПН"], "", "", "ЛОЖЬ", "ЛОЖЬ", "", "", record["Підстава"], record["Супровідний документ"], record["Текст наказу"]];
      }
      const joined = record["Дія"].startsWith("Зарахувати");
      return [record["Дія"], rank, record["ПІБ"], record["ІПН"], joined ? date : "", joined ? "" : date, joined ? "ІСТИНА" : "ЛОЖЬ", joined ? "ІСТИНА" : "ЛОЖЬ", joined ? foodDate : "", joined ? "" : foodDate, record["Підстава"], record["Супровідний документ"], record["Текст наказу"]];
    }
    const arrival = record["Дія"].startsWith("Прибув") || record["Дія"].startsWith("Поновити") || record["Дія"].includes("Закрити");
    const orderNumber = orderNumberText;
    if (record["Дія"].split(" | ").includes("Повернення до іншої частини")) {
      return [record["Дія"], rank, record["ПІБ"], record["ІПН"], orderNumber, record["Вибуття.Продовольче"], record["Вибуття"], "", "", "", record["Прибуття"], record["Прибуття.Продовольче"], orderNumber, record["Підстава"], record["Супровідний документ"], record["Підстава"], record["Текст наказу"]];
    }
    return [record["Дія"], rank, record["ПІБ"], record["ІПН"], ...(arrival ? ["", "", "", "", "", "", date, foodDate, orderNumber, "", record["Супровідний документ"], record["Підстава"], record["Текст наказу"]] : [orderNumber, foodDate, date, record["Термін"], record["Дорога"], record["Прибуття.План"] || "", "", "", "", record["Підстава"], record["Супровідний документ"], "", record["Текст наказу"]])];
  });
  if (name === "Відсутні") {
    values.forEach((row, index) => {
      const record = recordsForSheet[index];
      row.splice(4, 0, record["Посада"] ?? "", record["Подія"] ?? "", record["Подія"] === "Самовільне залишення частини" ? "" : record["Куди"] ?? "", "");
    });
  }
  if (uidIndex) values.forEach(row => row.splice(uidIndex, 0, ""));
  if (name === "Прибули") values.forEach(row => row.splice(3, 0, "", ""));
  if (name === "Прибули") values.forEach((row, index) => row.splice(row.length - 1, 0, recordsForSheet[index]["Додаткова інформація"] ?? ""));
  values = values.map((row, index) => selected.map(header => header === "Додаткова інформація" ? recordsForSheet[index][header] ?? "" : row[legacyHeaders.indexOf(header)] ?? ""));
  // Даты остаются числовыми значениями Excel, а ИПН — текстовыми идентификаторами.
  const dateHeaders = new Set(["Зарахування", "Виключення", "Вибуття.ПЗ", "Вибуття", "Прибуття.План", "Прибуття", "Прибуття.ПЗ", "Вибуття.План"]);
  for (const row of values) {
    outputHeaders.forEach((header, index) => {
      if (["Вибуття.Термін", "Вибуття.Дорога", "Прибуття.Термін"].includes(header) && /^\d+$/.test(row[index] ?? "")) {
        row[index] = Number(row[index]);
      }
      if (dateHeaders.has(header) && /^\d{2}\.\d{2}\.\d{4}$/.test(row[index] ?? "")) {
        const [day, month, year] = row[index].split(".").map(Number);
        row[index] = (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000;
      }
    });
  }
  if (values.length) sheet.getRangeByIndexes(4, 0, values.length, outputHeaders.length).values = values;
  const bottom = Math.max(5, values.length + 4);
  sheet.getRange(`A4:${lastCol}${bottom}`).format.font = { name: "Times New Roman", size: 10, color: "#FFFFFF" };
  sheet.getRange(`A1:${lastCol}1`).format = { fill: palette[name], font: { name: "Times New Roman", size: 14, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center" };
  sheet.getRange(`A2:${lastCol}2`).format = { font: { name: "Times New Roman", size: 9, italic: true, color: "#D9D9D9" } };
  sheet.getRange(`A4:${lastCol}4`).format = { fill: "#000000", font: { name: "Times New Roman", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
  sheet.getRange(`A5:${lastCol}${bottom}`).format.fill = "#262626";
  sheet.getRange(`A4:${lastCol}${bottom}`).format.borders = { preset: "all", style: "thin", color: "#0D0D0D" };
  sheet.getRange(`A4:${lastCol}${bottom}`).format.verticalAlignment = "center";
  sheet.getRange(`B5:B${bottom}`).format.horizontalAlignment = "center";
  sheet.getRange(`D5:D${bottom}`).format.horizontalAlignment = "center";
  sheet.getRange(`A4:${lastCol}${bottom}`).format.wrapText = true;
  outputHeaders.forEach((header, index) => {
    if (dateHeaders.has(header)) sheet.getRangeByIndexes(4, index, bottom - 4, 1).setNumberFormat("dd.mm.yyyy");
  });
  applyLayout(sheet, outputHeaders, values);
  sheet.getRange("A1:AZ1000").format.fill = "#28251E";
  sheet.getRange(`A1:${lastCol}1`).format.fill = "#5D5142";
  sheet.getRange(`A4:${lastCol}4`).format.fill = "#5D5142";
  for (let index = 0; index < values.length; index++) sheet.getRangeByIndexes(index + 4, 0, 1, outputHeaders.length).format.fill = index % 2 ? "#38342C" : "#302D25";
  sheet.getRange("A1").format.rowHeight = 25;
  if (values.length) {
    const table = sheet.tables.add(`A4:${lastCol}${bottom}`, true, tableNames[name]);
    table.style = "TableStyleDark1";
    table.showFilterButton = true;
  }
  // Высоту рассчитываем по переносам, не изменяя согласованные ширины колонок.
  for (let index = 0; index < values.length; index += 1) {
    const lineCount = Math.max(...values[index].map((value, column) => {
      if (value == null || value === "") return 1;
      const width = sheet.getRangeByIndexes(4 + index, column, 1, 1).format.columnWidth;
      return Math.ceil(String(value).length / Math.max(6, Math.floor(width * 1.15)));
    }));
    sheet.getRangeByIndexes(4 + index, 0, 1, outputHeaders.length).format.rowHeight = Math.max(30, Math.min(409, lineCount * 12 + 8));
  }
}

const supplementaryRanges = addSupplementary(workbook, supplemental);
workbook.recalculate();
console.log((await workbook.inspect({ kind: "table", range: "Проблеми!A4:F10", include: "values,formulas", tableMaxRows: 7, tableMaxCols: 6 })).ndjson);
console.log((await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 50 }, summary: "scan" })).ndjson);
await fs.mkdir(outputDir, { recursive: true });
for (const name of order) {
  const count = records.filter(row => row["Лист"] === name).length;
  const end = name === "Список" ? "N" : name === "Відсутні" ? "V" : "T";
  supplementaryRanges[name] = `A1:${end}${Math.min(count + 4, 12)}`;
}
for (const [sheetName, range] of Object.entries(supplementaryRanges)) {
  const preview = await workbook.render({ sheetName, range, scale: 0.8, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}
const output = await SpreadsheetFile.exportXlsx(workbook);
await fs.mkdir(dirname(outputPath), { recursive: true });
await output.save(outputPath);
// Масштаб сохраняется в представлениях всех листов готового XLSX.
const viewResult = spawnSync("C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe", [fileURLToPath(new URL("set_workbook_zoom.py", import.meta.url)), outputPath], { encoding: "utf8" });
if (viewResult.status !== 0) throw new Error(viewResult.stderr);