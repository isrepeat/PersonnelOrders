import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const args = process.argv.slice(2);
const folder = args[args.indexOf('--date') + 1];
const output = args[args.indexOf('--output') + 1];
if (!args.includes('--date') || !args.includes('--output')) throw new Error('Нужны --date YYYY.MM.DD и --output файл.xlsx');
const tools = path.dirname(fileURLToPath(import.meta.url));
const python = process.env.FOOD_PYTHON || 'C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';
const run = JSON.parse(execFileSync(python, [path.join(tools, 'process_food_reports.py'), '--date', folder], { encoding: 'utf8' }));
const wb = Workbook.create();
const serial = text => (Date.parse(text + 'T00:00:00Z') - Date.UTC(1899, 11, 30)) / 86400000;
const food = [], dry = [];
for (const group of run.groups) {
  for (const person of group.people) {
    const extra = [group.source_pages.join('; '), person.match, group.reporter, person.ipn || ''];
    if (group.kind === 'dry') dry.push([person.rank, person.name, group.military, serial(group.start), group.duration, null, group.basis, run.order.number, null, ...extra]);
    else food.push([person.rank, person.name, group.military, null, null, null, serial(group.start), run.order.number, group.basis, null, ...extra]);
  }
}
for (const [name, headers, values, dates] of [
  ['Продовольче', ['Звання', 'ПІБ', 'Установа', 'Група', 'Зарахування', 'Зарахування.Наказ', 'Виключення', 'Виключення.Наказ', 'Підстава', 'Додаткова інформація', 'Джерело', 'Звірка ШПО', 'Рапорт', 'ІПН'], food, ['F', 'H']],
  ['Сухпрод', ['Звання', 'ПІБ', 'ІПН/Установа', 'Початок', 'Тривалість', 'Припинення', 'Підстава', 'Наказ', 'Додаткова інформація', 'Джерело', 'Звірка ШПО', 'Рапорт', 'ІПН'], dry, ['E', 'G']]
]) {
  const sheet = wb.worksheets.add(name);
  sheet.showGridLines = false;
  const end = String.fromCharCode(65 + headers.length);
  const last = Math.max(3, values.length + 2);
  sheet.getRange(`A1:AN${Math.max(50, last + 8)}`).format = { fill: '#262626', font: { name: 'Times New Roman', size: 12, color: '#FFFFFF' } };
  sheet.getRangeByIndexes(1, 1, 1, headers.length).values = [headers];
  if (values.length) sheet.getRangeByIndexes(2, 1, values.length, headers.length).values = values;
  const table = sheet.tables.add(`B2:${end}${last}`, true, name === 'Сухпрод' ? 'DryRations' : 'FoodReports');
  table.style = 'TableStyleDark1';
  table.showFilterButton = true;
  sheet.getRange(`B2:${end}${last}`).format = { wrapText: false, verticalAlignment: 'center', rowHeight: 22.5, font: { color: '#FFFFFF', name: 'Times New Roman', size: 12 }, borders: { preset: 'all', style: 'thin', color: '#000000' } };
  sheet.getRange(`B2:${end}2`).format = { fill: '#000000', font: { color: '#FFFFFF', bold: true }, horizontalAlignment: 'center', wrapText: true, rowHeight: 31.5 };
  for (let row = 3; row <= last; row++) sheet.getRange(`B${row}:${end}${row}`).format.fill = row % 2 ? '#262626' : '#383838';
  const widths = name === 'Сухпрод' ? [8.625, 14.75, 40.375, 17.75, 13.375, 12.5, 16.25, 55.875, 8.5, 36.5] : [8.625, 21.25, 42, 13.25, 12.625, 13.375, 13, 13.75, 13.25, 58, 25.25];
  widths.forEach((width, index) => sheet.getRange(`${String.fromCharCode(65 + index)}:${String.fromCharCode(65 + index)}`).format.columnWidth = width);
  sheet.getRange(name === 'Сухпрод' ? 'K:N' : 'L:O').format.columnWidth = 28;
  if (values.length) {
    sheet.getRange(`D3:D${last}`).format = { font: { color: '#FF9966' }, horizontalAlignment: 'center' };
    sheet.getRange(`E3:I${last}`).format.horizontalAlignment = 'center';
    sheet.getRange(name === 'Сухпрод' ? `H3:H${last}` : `J3:J${last}`).format.horizontalAlignment = 'left';
    if (name === 'Продовольче') sheet.getRange(`J3:J${last}`).format = { wrapText: true, rowHeight: 45 };
    for (let row = 3; row <= last; row++) if (values[row - 3][0].length > 13) sheet.getRange(`B${row}`).format = { wrapText: true, rowHeight: 31.5 };
    for (const col of dates) sheet.getRange(`${col}3:${col}${last}`).setNumberFormat('dd.mm.yyyy');
    if (name === 'Сухпрод') {
      sheet.getRange(`G3:G${last}`).format.font.color = '#8DB4E2';
      for (let row = 3; row <= last; row++) sheet.getRange(`G${row}`).formulas = [['=[@Початок]+[@Тривалість]']];
    }
    sheet.getRange(`${end}3:${end}${last}`).setNumberFormat('@');
    sheet.getRange(`C3:C${last}`).conditionalFormats.add('duplicateValues', { format: { fill: '#FFC7CE', font: { color: '#9C0006' } } });
  }
  sheet.freezePanes.unfreeze();
}
wb.recalculate();
console.log((await wb.inspect({ kind: 'table', range: 'Продовольче!B2:K3', include: 'values,formulas', tableMaxRows: 2, tableMaxCols: 10, maxChars: 1000 })).ndjson);
console.log((await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!', options: { useRegex: true, maxResults: 10 }, summary: 'Ошибки формул Food' })).ndjson);
await fs.mkdir(path.dirname(output), { recursive: true });
await (await SpreadsheetFile.exportXlsx(wb)).save(output);
for (const name of ['Продовольче', 'Сухпрод']) {
  const preview = await wb.render({ sheetName: name, range: name === 'Сухпрод' ? 'A1:K7' : 'A1:L4', scale: 1.3, format: 'png' });
  await fs.writeFile(`${run.work}/${name}.png`, new Uint8Array(await preview.arrayBuffer()));
}
console.log(JSON.stringify({ folder, order: run.order.number, food: food.length, dry: dry.length, output }));