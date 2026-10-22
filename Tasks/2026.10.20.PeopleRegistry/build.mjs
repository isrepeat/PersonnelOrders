import fs from 'node:fs/promises';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const data = JSON.parse(await fs.readFile(new URL('./source.json', import.meta.url), 'utf8'));
const norm = value => String(value ?? '').replace(/[\u00a0\u202f]/g, ' ').replace(/\s+/g, ' ').trim();
const key = value => norm(value).toLocaleUpperCase('uk-UA');
const isTax = value => /^[0-9]{10}$/.test(norm(value));
const token = value => norm(value).replace(/[ _.,;:/\\|\[\](){}<>!?"'`]/g, '');
const dateParts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
  timeZone: 'Europe/Kyiv', year: 'numeric', month: '2-digit', day: '2-digit',
  hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23'
}).formatToParts(new Date()).map(part => [part.type, part.value]));
const timestamp = `${dateParts.year}-${dateParts.month}-${dateParts.day} ${dateParts.hour}:${dateParts.minute}:${dateParts.second}`;
const listNames = new Map();
const idOwners = new Map();
const sourceIDs = new Set();
for (const row of [...data.list, ...data.arrivals]) {
  if (norm(row.id)) {
    sourceIDs.add(key(row.id));
    if (!idOwners.has(key(row.id))) idOwners.set(key(row.id), new Set());
    idOwners.get(key(row.id)).add(key(row.name));
  }
}
for (const row of data.list) {
  if (!norm(row.name) || !norm(row.id)) continue;
  if (!listNames.has(key(row.name))) listNames.set(key(row.name), new Map());
  const candidates = listNames.get(key(row.name));
  if (!candidates.has(key(row.id))) candidates.set(key(row.id), { id: norm(row.id), rows: [] });
  candidates.get(key(row.id)).rows.push(row.row);
}
const arrivalsByName = new Map();
for (const row of data.arrivals) {
  if (!arrivalsByName.has(key(row.name))) arrivalsByName.set(key(row.name), []);
  arrivalsByName.get(key(row.name)).push(row.row);
}
const historyIDs = new Set(data.reservedHistoryIDs.map(key));
const historyPeople = new Map();
for (const row of data.historyRecords) {
  const match = key(row.name) + '\u001f' + key(row.origin);
  if (!historyPeople.has(match)) historyPeople.set(match, new Set());
  historyPeople.get(match).add(row.currentID);
}
const reserved = new Set([...sourceIDs, ...historyIDs]);
function generateID(name, origin) {
  const words = norm(name).split(' ');
  const surname = token(words[0]);
  let base = surname.charAt(0).toLocaleUpperCase('uk-UA') + surname.slice(1).toLocaleLowerCase('uk-UA');
  base += words.slice(1).map(word => token(word).charAt(0).toLocaleUpperCase('uk-UA')).join('');
  base += '_' + token(origin);
  base = base.slice(0, 80);
  let candidate = base;
  let counter = 1;
  while (reserved.has(key(candidate))) {
    const suffix = '_' + (++counter);
    candidate = base.slice(0, 80 - suffix.length) + suffix;
  }
  reserved.add(key(candidate));
  return candidate;
}

const rows = [];
const imports = new Map();
const stats = { rows: data.arrivals.length, tax: 0, existing: 0, temporary: 0, ambiguous: 0, repeats: 0, missing: 0, historySkipped: 0 };
const issued = new Set();
for (const row of data.arrivals) {
  const name = norm(row.name);
  const origin = norm(row.origin);
  const existing = norm(row.id);
  const candidates = [...(listNames.get(key(name))?.values() ?? [])];
  const savedIDs = [...(historyPeople.get(key(name) + '\u001f' + key(origin)) ?? [])];
  let id = '';
  let type = '';
  let status = 'Готово';
  let note = '';
  let ready = true;
  if (!name) {
    status = 'Перевірити ПІБ'; note = 'У джерелі немає ПІБ.'; stats.missing++; ready = false;
  } else if (candidates.length > 1) {
    status = 'Перевірити збіг'; note = 'У Список для цього ПІБ є кілька різних ID. Автоматичний вибір не виконано.';
    stats.ambiguous++; ready = false;
  } else if (candidates.length === 1) {
    id = candidates[0].id;
    type = isTax(id) ? 'ІПН зі Список' : 'ID зі Список';
    if (idOwners.get(key(id))?.size > 1) {
      status = 'Перевірити власника'; note = 'Наявний ID пов’язаний із різними ПІБ.'; ready = false; stats.ambiguous++;
    } else {
      note = 'Єдиний ID за точним ПІБ після нормалізації регістру та пробілів.';
      if (isTax(id)) stats.tax++; else stats.existing++;
    }
  } else if (existing) {
    id = existing;
    type = isTax(id) ? 'ІПН з Прибулі' : 'ID з Прибулі';
    if (idOwners.get(key(id))?.size > 1) {
      status = 'Перевірити власника'; note = 'Наявний ID пов’язаний із різними ПІБ.'; ready = false; stats.ambiguous++;
    } else {
      note = 'Збережено наявний ID з Прибулі; у Список збігу немає.';
      if (isTax(id)) stats.tax++; else stats.existing++;
    }
  } else if (savedIDs.length === 1) {
    id = savedIDs[0];
    type = 'ID з історії';
    note = 'Використано поточний ID зі збереженої історії за точними ПІБ та Звідки.';
    stats.existing++;
  } else if (savedIDs.length > 1) {
    status = 'Перевірити історію'; note = 'У збереженій історії є різні поточні ID для цих ПІБ та Звідки.';
    stats.ambiguous++; ready = false;
  } else if (!token(origin) || !token(name.split(' ')[0])) {
    status = 'Перевірити дані'; note = 'Для тимчасового ID потрібні ПІБ та Звідки.'; ready = false; stats.missing++;
  } else {
    id = generateID(name, origin);
    type = 'Тимчасовий';
    stats.temporary++;
    if (issued.has(key(id))) throw new Error('Duplicate generated ID');
    issued.add(key(id));
    if (arrivalsByName.get(key(name)).length > 1) {
      status = 'Перевірити повтор';
      note = 'ПІБ повторюється у Прибулі без однозначного ID зі Список. ID для рядка — лише пропозиція; особу не об’єднано.';
      stats.repeats++; ready = false;
    } else {
      note = 'Новий ID-пропозиція; не зайнятий у збережених РУХ та історії на час підготовки.';
    }
  }
  if (existing && id && key(existing) !== key(id)) note += ' Наявний ID у Прибулі відрізняється; перевірте заміну.';
  if (id && savedIDs.length && !savedIDs.some(savedID => key(savedID) === key(id))) {
    status = 'Перевірити заміну';
    note += ' У реєстрі є інший поточний ID; заміну треба виконати зі збереженням зв’язку через форму.';
    ready = false;
  }
  if (id && historyIDs.has(key(id))) {
    note += ' ID уже присутній у збереженій історії; повторно не імпортувати.';
    stats.historySkipped++;
  }
  rows.push([name, origin, id, type, row.row, existing, status, note, candidates.map(c => c.id).join(' | ')]);
  if (ready && id && !historyIDs.has(key(id)) && !imports.has(key(id))) imports.set(key(id), [id, id, name, origin, timestamp]);
  if (ready && id && existing && key(existing) !== key(id) && !historyIDs.has(key(existing))) {
    if (!imports.has(key(existing))) imports.set(key(existing), [existing, id, name, origin, timestamp]);
  }
}
stats.importRows = imports.size;
stats.reviewRows = rows.filter(row => row[6] !== 'Готово').length;
const importRows = [...imports.values()];
if (rows.length !== data.arrivals.length) throw new Error('Lost arrival rows');
for (const row of importRows) {
  const current = imports.get(key(row[1]));
  if ((!current || key(current[0]) !== key(current[1])) && !historyIDs.has(key(row[1]))) throw new Error('Invalid current-ID link');
}
const workbook = Workbook.create();
const sheet = workbook.worksheets.add('Прибулі — ID');
const history = workbook.worksheets.add('Для реєстру');
const colors = { background: '#202124', alternate: '#303134', foreground: '#E8EAED', header: '#235B70' };

function styleSheet(target, headers, values, widths, tableName, title, summary, note) {
  const lastCol = String.fromCharCode(64 + headers.length);
  const lastRow = values.length + 6;
  target.showGridLines = false;
  target.tabColor = colors.header;
  const area = target.getRange(`A1:${lastCol}${lastRow + 2}`);
  area.format.fill = colors.background;
  area.format.font = { name: 'Arial', size: 10, color: colors.foreground };
  area.format.rowHeight = 34;
  area.format.verticalAlignment = 'center';
  area.format.wrapText = true;
  target.getRange(`A1:${lastCol}1`).merge();
  target.getRange('A1').values = [[title]];
  target.getRange('A1').format.font = { size: 16, bold: true, color: colors.foreground };
  target.getRange(`A2:${lastCol}2`).merge();
  target.getRange('A2').values = [[summary]];
  target.getRange(`A3:${lastCol}3`).merge();
  target.getRange('A3').values = [[`Джерело: ${data.path} · збережена версія · підготовлено ${timestamp}`]];
  target.getRange(`A4:${lastCol}4`).merge();
  target.getRange('A4').values = [[note]];
  target.getRange(`A4:${lastCol}4`).format.rowHeight = 42;
  target.getRange(`A6:${lastCol}${lastRow}`).values = [headers, ...values];
  const table = target.tables.add(`A6:${lastCol}${lastRow}`, true, tableName);
  table.style = 'TableStyleDark1';
  table.showFilterButton = true;
  target.getRange(`A6:${lastCol}6`).format.fill = colors.header;
  target.getRange(`A6:${lastCol}6`).format.font = { bold: true, color: colors.foreground };
  target.getRange(`A6:${lastCol}6`).format.rowHeight = 30;
  target.getRange(`A6:${lastCol}6`).format.horizontalAlignment = 'center';
  for (let index = 0; index < values.length; index++) {
    target.getRange(`A${index + 7}:${lastCol}${index + 7}`).format.fill = index % 2 ? colors.alternate : colors.background;
  }
  widths.forEach((width, index) => target.getRange(`${String.fromCharCode(65 + index)}1:${String.fromCharCode(65 + index)}${lastRow}`).format.columnWidth = width);
  target.freezePanes.freezeRows(6);
  target.freezePanes.freezeColumns(1);
}

styleSheet(sheet, ['ПІБ', 'Звідки', 'ID для перенесення', 'Тип ID', 'Рядок Прибулі', 'Наявний ID у Прибулі', 'Статус', 'Примітка', 'ID зі Список'], rows,
  [36, 40, 56, 20, 12, 24, 25, 78, 30], 'tbArrivalIDProposals', 'Ідентифікатори для Прибулі',
  `Рядків: ${stats.rows} · ІПН: ${stats.tax} · наявні ID: ${stats.existing} · тимчасові пропозиції: ${stats.temporary} · на перевірку: ${stats.reviewRows}`,
  'Колонка C — результат для відповідного рядка РУХ (номер у E). Статус Перевірити означає: не переносити без ручної перевірки. Нові ID ще не зарезервовані.');
sheet.getRange(`C7:D${rows.length + 6}`).setNumberFormat('@');
sheet.getRange(`F7:F${rows.length + 6}`).setNumberFormat('@');
sheet.getRange(`I7:I${rows.length + 6}`).setNumberFormat('@');
sheet.getRange(`G7:G${rows.length + 6}`).conditionalFormats.add('containsText', { text: 'Перевірити', format: { fill: '#5B4320', font: { color: '#FFD590', bold: true } } });
styleSheet(history, ['ID', 'Поточний ID', 'ПІБ', 'Звідки', 'Дата видачі'], importRows,
  [53, 53, 38, 42, 23], 'tbPreparedIDHistory', 'Підготовлено для історії ID',
  `Унікальних ID для перенесення: ${stats.importRows}. Спірні рядки та ID, уже наявні у збереженій історії, не включено.`,
  'Порядок колонок відповідає історії реєстру. Дата — час підготовки пропозиції, не дата отримання ІПН. Перед імпортом перевірте ФІО-збіги й актуальну зайнятість ID.');
history.getRange(`A7:E${importRows.length + 6}`).setNumberFormat('@');
workbook.recalculate();
const check = await workbook.inspect({ kind: 'table', range: 'Прибулі — ID!A6:G10', include: 'values', tableMaxRows: 5, tableMaxCols: 7 });
await fs.writeFile(new URL('./inspection.json', import.meta.url), check.ndjson);
const preview = await workbook.render({ sheetName: 'Прибулі — ID', range: 'A1:I12', scale: 1.2, format: 'png' });
await fs.writeFile(new URL('./preview.png', import.meta.url), new Uint8Array(await preview.arrayBuffer()));
const previewHistory = await workbook.render({ sheetName: 'Для реєстру', range: 'A1:E11', scale: 1.3, format: 'png' });
await fs.writeFile(new URL('./preview-history.png', import.meta.url), new Uint8Array(await previewHistory.arrayBuffer()));
const digest = crypto.createHash('sha256').update(await fs.readFile(data.path)).digest('hex');
if (digest !== data.sha256) throw new Error('RUH changed during preparation; regenerate from the latest saved source.');
await fs.mkdir(new URL('./outputs/', import.meta.url), { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(fileURLToPath(new URL('./outputs/Прибулі_ID_2026-10-20.xlsx', import.meta.url)));
await fs.writeFile(new URL('./stats.json', import.meta.url), JSON.stringify(stats, null, 2));
console.log(JSON.stringify(stats));