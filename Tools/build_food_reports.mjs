import fs from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const [input, output, work] = process.argv.slice(2);
const groups = JSON.parse(await fs.readFile(input, 'utf8'));
// Читаем номера приказов из tbOrders по дате рапорта, независимо от начала сухпайка.
const python = process.env.FOOD_PYTHON || 'C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';
const orderNumbers = JSON.parse(execFileSync(python, ['Tools/fill_food_order_numbers.py', '--mapping'], { encoding: 'utf8' }));
const wb = Workbook.create();
const serial = text => (Date.parse(text + 'T00:00:00Z') - Date.UTC(1899, 11, 30)) / 86400000;
const food = [], dry = [];
for (const g of groups) {
  for (const p of g.people) {
    const extra = [(g.source_pages || [g.source]).join('; '), p.match, g.reporter, p.ipn || ''];
    const reporterPosition = { КОРНЕНКО: 'командира зведеного підрозділу мобільного зв’язку 1 ОПЗ військової частини А0565', МОРОЗОВИЧ: 'командира роти радіоелектронної боротьби', МІЗЯК: 'командира 1 батальйону безпілотних систем' }[g.reporter];
    const basis = { КОРНЕНКО: 'рапорт майора Юрія КОРНЕНКА (вх. № 1656/26791-в від 08.10.2026)', МОРОЗОВИЧ: 'рапорт лейтенанта МОРОЗОВИЧА\u00a0Ю.М. (вх. № 1656/26786-в від 08.10.2026)', МІЗЯК: 'рапорт капітана МІЗЯКА\u00a0А.В. (вх. № 1656/26790-в від 08.10.2026)' }[g.reporter].replace(/^рапорт /, `рапорт ${reporterPosition} `);
    dry.push([p.rank, p.name, g.military, serial(g.start), g.duration, null, basis, orderNumbers[g.report_date || '2026-10-08'] ?? null, null, ...extra]);
  }
}
for (const [name, headers, values, dates] of [
  ['Продовольче', ['Звання', 'ПІБ', 'Установа', 'Група', 'Зарахування', 'Зарахування.Наказ', 'Виключення', 'Виключення.Наказ', 'Підстава', 'Додаткова інформація', 'Джерело', 'Звірка ШПО', 'Рапорт', 'ІПН'], food, ['F', 'H']],
  ['Сухпрод', ['Звання', 'ПІБ', 'ІПН/Установа', 'Початок', 'Тривалість', 'Припинення', 'Підстава', 'Наказ', 'Додаткова інформація', 'Джерело', 'Звірка ШПО', 'Рапорт', 'ІПН'], dry, ['E', 'G']]
]) {
  const s = wb.worksheets.add(name);
  s.showGridLines = false;
  s.getRange('A1:AN300').format = { fill: '#262626', font: { name: 'Times New Roman', size: 12, color: '#FFFFFF' } };
  s.getRangeByIndexes(1, 1, 1, headers.length).values = [headers];
  if (values.length) s.getRangeByIndexes(2, 1, values.length, headers.length).values = values;
  const end = String.fromCharCode(65 + headers.length);
  s.tables.add(`B2:${end}${Math.max(3, values.length + 2)}`, true, name === 'Сухпрод' ? 'DryRations' : 'FoodReports').style = 'TableStyleDark1';
  s.getRange(`B2:${end}${values.length + 2}`).format = { wrapText: false, verticalAlignment: 'center', rowHeight: 15.75, font: { color: '#FFFFFF', name: 'Times New Roman', size: 12 }, borders: { preset: 'all', style: 'thin', color: '#000000' } };
  s.getRange(`B2:${end}2`).format = { fill: '#000000', font: { color: '#FFFFFF', bold: true }, horizontalAlignment: 'center', rowHeight: 26.25 };
  for (let n = 3; n <= values.length + 2; n++) s.getRange(`B${n}:${end}${n}`).format.fill = n % 2 ? '#262626' : '#383838';
  // Ширина восстановлена по границам колонок на скриншоте, в пикселях при масштабе 100%.
  const widths = name === 'Сухпрод' ? [69, 118, 323, 143, 107, 100, 129, 448, 68, 293] : [69, 118, 323, 143, 107, 129, 129, 129, 129, 448, 293];
  widths.forEach((pixels, index) => s.getRange(`${String.fromCharCode(65 + index)}:${String.fromCharCode(65 + index)}`).format.columnWidth = (pixels - 5) / 7);
  s.getRange(name === 'Сухпрод' ? 'K:N' : 'L:O').format.columnWidth = 28;
  if (!values.length) { s.freezePanes.unfreeze(); continue; }
  s.getRange(`D3:D${values.length + 2}`).format = { font: { color: '#FF9966' }, horizontalAlignment: 'center' };
  s.getRange(`E3:I${values.length + 2}`).format.horizontalAlignment = 'center';
  s.getRange(name === 'Сухпрод' ? `H3:H${values.length + 2}` : `J3:J${values.length + 2}`).format.horizontalAlignment = 'left';
  for (let n = 3; n <= values.length + 2; n++) if (values[n - 3][0].length > 13) {
    s.getRange(`B${n}`).format.wrapText = true;
    s.getRange(`B${n}`).format.rowHeight = 31.5;
  }
  if (name === 'Сухпрод') s.getRange(`G3:G${values.length + 2}`).format.font.color = '#8DB4E2';
  for (const col of dates) s.getRange(`${col}3:${col}${values.length + 2}`).setNumberFormat('dd.mm.yyyy');
  if (name === 'Сухпрод') for (let n = 3; n <= values.length + 2; n++) s.getRange(`G${n}`).formulas = [['=[@Початок]+[@Тривалість]']];
  s.freezePanes.unfreeze();
}
wb.recalculate();
await (await SpreadsheetFile.exportXlsx(wb)).save(output);
for (const name of ['Продовольче', 'Сухпрод']) {
  const preview = await wb.render({ sheetName: name, range: 'A1:K7', scale: 1.3, format: 'png' });
  await fs.writeFile(`${work}/${name}.png`, new Uint8Array(await preview.arrayBuffer()));
}
console.log(JSON.stringify({food: food.length, dry: dry.length}));