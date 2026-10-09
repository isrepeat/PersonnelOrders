import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

// Сборщик принимает проверенные данные и явный путь новой версии, не перезаписывает старую.
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
if (payload) {
  payload.notes ??= [];
  payload.food ??= [];
  for (const record of records) {
    if ((record["Деталі"] || "").startsWith("ТВО:")) payload.notes.push(`${record["ПІБ"]}: ${record["Деталі"]}`);
  }
  payload.notes = [...new Set(payload.notes)];
}
const orderNumberText = records[0]["Наказ"].match(/№(\d+)/)[1];
const orderDateText = records[0]["Наказ"].match(/\d{2}\.\d{2}\.\d{4}/)[0];
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
  const absenceHeaders = ["Дія", "Звання", "ПІБ", "ІПН", "Посада", "Подія", "Куди", "За межі", "Вибуття.Наказ", "Вибуття.Продовольче", "Вибуття", "Вибуття.Термін", "Вибуття.Дорога", "Прибуття.План", "Прибуття", "Прибуття.Продовольче", "Прибуття.Наказ", "Вибуття.Підстава", "Супровідний документ", "Прибуття.Підстава", "ТВО.ПІБ", "ТВО.ІПН", "ТВО.Посада", "Текст наказу"];
  const arrivalHeaders = ["Дія", "Звання", "ПІБ", "Посада", "Звідки", "Прибуття.Наказ", "Прибуття.Продовольче", "Прибуття", "Прибуття.Термін", "Вибуття.План", "Вибуття", "Вибуття.Продовольче", "Вибуття.Наказ", "Прибуття.Підстава", "Супровідний документ", "Вибуття.Підстава", "Текст наказу"];
  const rosterHeaders = ["Дія", "Звання", "ПІБ", "ІПН", "Зарахування", "Виключення", "В строю", "На продовольчому", "Зараховано на продовольче", "Виключено з продовольчого", "Підстава", "Супровідний документ", "Текст наказу"];
  const outputHeaders = name === "Відсутні" ? absenceHeaders : name === "Прибули" ? arrivalHeaders : rosterHeaders;
  const lastCol = String.fromCharCode(64 + outputHeaders.length);
  sheet.getRange(`A1:${lastCol}1`).merge();
  sheet.getRange("A1").values = [[`${name}: події для ручного внесення`]];
  sheet.getRange(`A2:${lastCol}2`).merge();
  sheet.getRange("A2").values = [[`Джерело: наказ №${orderNumberText} від ${orderDateText}. ПІБ звірено з ШПО за ІПН; без ІПН приведено до називного відмінка.`]];
  sheet.getRange(`A4:${lastCol}4`).values = [outputHeaders];
  const values = recordsForSheet.map((record) => {
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
      row.splice(4, 0, record["Посада"] ?? "", record["Подія"] ?? "", record["Куди"] ?? "", "");
      row.splice(row.length - 1, 0, record["ТВО.ПІБ"] ?? "", record["ТВО.ІПН"] ?? "", record["ТВО.Посада"] ?? "");
    });
  }
  // Даты остаются числовыми значениями Excel, а ИПН — текстовыми идентификаторами.
  const dateHeaders = new Set(["Зарахування", "Виключення", "Зараховано на продовольче", "Виключено з продовольчого", "Вибуття.Продовольче", "Вибуття", "Прибуття.План", "Прибуття", "Прибуття.Продовольче", "Вибуття.План"]);
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
    if (dateHeaders.has(header)) sheet.getRangeByIndexes(4, index, bottom - 4, 1).format.numberFormat = "dd.mm.yyyy";
  });
  sheet.getRange(`B4:${lastCol}${bottom}`).format.horizontalAlignment = "center";
  sheet.getRange(`A4:A${bottom}`).format.horizontalAlignment = "left";
  sheet.getRange(`C4:C${bottom}`).format.horizontalAlignment = "left";
  if (name === "Відсутні") {
    sheet.getRange(`R4:R${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`T4:T${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`X4:X${bottom}`).format.horizontalAlignment = "left";
  }
  if (name === "Прибули") {
    sheet.getRange(`A4:A${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`C4:C${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`N4:N${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`P4:P${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`Q4:Q${bottom}`).format.horizontalAlignment = "left";
  }
  if (name === "Список") {
    sheet.getRange(`A4:A${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`C4:C${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`K4:K${bottom}`).format.horizontalAlignment = "left";
    sheet.getRange(`M4:M${bottom}`).format.horizontalAlignment = "left";
  }
  sheet.getRange("A:A").format.columnWidth = 15;
  sheet.getRange("B:B").format.columnWidth = 10;
  sheet.getRange("C:C").format.columnWidth = 28;
  sheet.getRange("D:D").format.columnWidth = 14;
  sheet.getRange("E:E").format.columnWidth = 14;
  sheet.getRange("F:F").format.columnWidth = 19;
  sheet.getRange("G:G").format.columnWidth = 34;
  sheet.getRange("H:H").format.columnWidth = 68;
  if (name === "Відсутні") {
    sheet.getRange("E:E").format.columnWidth = 14;
    sheet.getRange("F:F").format.columnWidth = 23;
    sheet.getRange("G:G").format.columnWidth = 34;
    sheet.getRange("H:H").format.columnWidth = 8;
    sheet.getRange("I:I").format.columnWidth = 7;
    sheet.getRange("J:J").format.columnWidth = 12;
    sheet.getRange("K:K").format.columnWidth = 10;
    sheet.getRange("L:L").format.columnWidth = 10;
    sheet.getRange("M:M").format.columnWidth = 10;
    sheet.getRange("N:N").format.columnWidth = 10;
    sheet.getRange("O:O").format.columnWidth = 10;
    sheet.getRange("P:P").format.columnWidth = 12;
    sheet.getRange("Q:Q").format.columnWidth = 9;
    sheet.getRange("R:R").format.columnWidth = 34;
    sheet.getRange("S:S").format.columnWidth = 14;
    sheet.getRange("T:T").format.columnWidth = 34;
    sheet.getRange("U:U").format.columnWidth = 28;
    sheet.getRange("V:V").format.columnWidth = 14;
    sheet.getRange("W:W").format.columnWidth = 14;
    sheet.getRange("X:X").format.columnWidth = 68;
    sheet.getRange(`V5:W${bottom}`).setNumberFormat("@");
    sheet.getRange(`E5:E${bottom}`).setNumberFormat("@");
  }
  if (name === "Прибули") {
    [7, 9, 10, 10, 10, 10, 10, 12, 9].forEach((width, index) => sheet.getRange(`${String.fromCharCode(70 + index)}:${String.fromCharCode(70 + index)}`).format.columnWidth = width);
    sheet.getRange("D:D").format.columnWidth = 28;
    sheet.getRange("E:E").format.columnWidth = 16;
    sheet.getRange("N:N").format.columnWidth = 34;
    sheet.getRange("O:O").format.columnWidth = 14;
    sheet.getRange("P:P").format.columnWidth = 34;
    sheet.getRange("Q:Q").format.columnWidth = 68;
  }
  if (name === "Список") {
    sheet.getRange("E:J").format.columnWidth = 14;
    sheet.getRange("K:K").format.columnWidth = 34;
    sheet.getRange("L:L").format.columnWidth = 14;
    sheet.getRange("M:M").format.columnWidth = 68;
  }
  sheet.getRange("A1").format.rowHeight = 25;
  sheet.getRange(`A4:${lastCol}4`).format.rowHeight = 30;
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

const overview = workbook.worksheets.add("Огляд");
overview.showGridLines = false;
overview.tabColor = "#404040";
overview.getRange("A1:Z1000").format.fill = "#262626";
overview.getRange("A1:D1").merge();
overview.getRange("A1").values = [[`Рух особового складу: наказ №${orderNumberText}, витяг для перевірки`]];
overview.getRange("A3:D3").values = [["Лист", "Кількість подій", "Період наказів", "Призначення"]];
overview.getRange("A4:D6").values = order.map((name) => [name, records.filter((record) => record["Лист"] === name).length, `№${orderNumberText}, ${orderDateText}`, "Ручно звірити та внести у РУХ_last.xlsx"]);
overview.getRange("A8:D8").merge();
overview.getRange("A8").values = [["Перед перенесенням звірити події з актуальною цільовою книгою та не дублювати наявні записи."]];
overview.getRange("A8:D8").format = { wrapText: true, verticalAlignment: "center", font: { name: "Times New Roman", size: 10, color: "#FFFFFF" }, rowHeight: 65 };
overview.getRange("A10:D10").merge();
overview.getRange("A10").values = [[records.filter(r => (r["Деталі"] || "").startsWith("ТВО:")).map(r => `${r["ПІБ"]}: ${r["Деталі"]}`).join("\n")]];
overview.getRange("A11:D11").merge();
overview.getRange("A11").values = [[""]];
overview.getRange("A10:D11").format = { wrapText: true, verticalAlignment: "center", font: { name: "Times New Roman", size: 10, color: "#FFFFFF" }, rowHeight: 65 };
overview.getRange("A1:D1").format = { fill: "#404040", font: { name: "Arial", size: 14, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center" };
overview.getRange("A3:D3").format = { fill: "#1F1F1F", font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
overview.getRange("A3:D6").format = { font: { name: "Times New Roman", size: 10, color: "#FFFFFF" }, borders: { preset: "all", style: "thin", color: "#0D0D0D" } };
overview.getRange("A4:D6").format.fill = "#262626";
overview.getRange("A:A").format.columnWidth = 18;
overview.getRange("B:B").format.columnWidth = 18;
overview.getRange("C:C").format.columnWidth = 30;
overview.getRange("D:D").format.columnWidth = 40;
overview.getRange("A1").format.rowHeight = 25;
overview.getRange("A1:D11").format.wrapText = true;
overview.getRange("A1:D11").format.verticalAlignment = "center";
if (payload) {
  overview.getRange('A8:D11').unmerge();
  overview.getRange('A8:D11').clear({applyTo: 'all'});
  overview.getRange('A8:F1000').format.fill = '#262626';
  for (let i=0; i<payload.notes.length; i++) {
    const n=i+8;
    overview.getRange(`A${n}:D${n}`).merge();
    overview.getRange(`A${n}`).values=[[payload.notes[i]]];
    overview.getRange(`A${n}:D${n}`).format={font:{name:'Times New Roman',size:10,color:'#FFFFFF'}, wrapText:true, verticalAlignment:'center', rowHeight:60};
  }
  if (payload.food.length) {
  const start=payload.notes.length+10;
  const foodHeaders=['ПІБ', 'Підрозділ', 'Зняти з котлового', 'По дату включно', 'Підстава', 'Текст наказу'];
  overview.getRange(`A${start}:F${start}`).values=[foodHeaders];
  if (payload.food.length) {
    const foodValues = payload.food.map(row => row.map((value, index) => {
      if ((index === 2 || index === 3) && /^\d{2}\.\d{2}\.\d{4}$/.test(value ?? "")) {
        const [day, month, year] = value.split(".").map(Number);
        return (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000;
      }
      return value;
    }));
    overview.getRangeByIndexes(start,0,foodValues.length,6).values=foodValues;
    overview.getRangeByIndexes(start,2,foodValues.length,2).setNumberFormat("dd.mm.yyyy");
  }
  const foodEnd=start+payload.food.length;
  overview.getRange(`A${start}:F${foodEnd}`).format={font:{name:'Times New Roman',size:10,color:'#FFFFFF'},wrapText:true,verticalAlignment:'center',rowHeight:90};
  overview.getRange(`A${start}:F${start}`).format.fill='#000000';
  overview.getRange('E:F').format.columnWidth=45;
  overview.tables.add(`A${start}:F${foodEnd}`,true,'DryRationReview').style='TableStyleDark1';
  }
}
workbook.recalculate();
const inspection = await workbook.inspect({ kind: "table", range: "Огляд!A1:D6", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 6 });
console.log(inspection.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 50 }, summary: "scan" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "Огляд", range: "A1:D11", scale: 2, format: "png" });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(`${outputDir}/overview.png`, new Uint8Array(await preview.arrayBuffer()));
for (const sheetName of ["Список", "Відсутні", "Прибули"]) {
  const sourceRows = records.filter((record) => record["Лист"] === sheetName).length;
  const endColumn = sheetName === "Список" ? "M" : sheetName === "Відсутні" ? "X" : "Q";
  const rendered = await workbook.render({ sheetName, range: `A1:${endColumn}${Math.min(sourceRows + 4, 12)}`, scale: 1.5, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName}.png`, new Uint8Array(await rendered.arrayBuffer()));
}
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
const closures = await workbook.render({ sheetName: "Відсутні", range: `A${Math.max(4, records.filter(r=>r['Лист']==='Відсутні').length)}:H${records.filter(r=>r['Лист']==='Відсутні').length+4}`, scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/closures.png`, new Uint8Array(await closures.arrayBuffer()));