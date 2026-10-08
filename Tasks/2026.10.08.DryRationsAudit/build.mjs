import fs from 'node:fs/promises';
import { gunzipSync } from 'node:zlib';
import { Workbook, SpreadsheetFile } from 'file:///C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';

const dir = 'C:/WORK/Windows/Строевые приказы/outputs/dry-rations-audit-20261008';
const data = JSON.parse(gunzipSync(await fs.readFile(`${dir}/reconciled.json.gz`)).toString('utf8'));
const assigned = data.groups.reduce((sum, g) => sum + g.people.length, 0);
const added = data.counts['Добавлено'] || 0;
const matched = assigned - added;
const uncertain = data.rows.filter(r => r.event?.sourceStart).length;
const dateDifferences = data.rows.filter(r => r.dateDifference).length;
const unconfirmed = data.counts['Не найдено в проверенных приказах'] || 0;
const latest = data.files.at(-1);
const latestDate = latest.slice(0, 10).split('-').reverse().join('.');
const latestOrder = latest.match(/№\s*(\d+)/)[1];
const monthNote = uncertain ? `№257: у ${uncertain} записей спорный месяц (август в приказе, сентябрь в таблице); дата таблицы сохранена и требует проверки.` : 'Спорных месяцев между сопоставленными назначениями и таблицей не обнаружено.';
const wb = Workbook.create();
const main = wb.worksheets.add('Сухпрод');
const audit = wb.worksheets.add('Сверка');
const orders = wb.worksheets.add('Пункты приказов');
const serial = x => typeof x === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(x) ? (Date.parse(`${x}T00:00:00Z`) - Date.UTC(1899, 11, 30)) / 86400000 : x;
const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

function layout(sheet, headers, rows, widths, title, notes, tableName) {
  const lastCol = letters[headers.length - 1];
  const last = rows.length + 10;
  sheet.showGridLines = false;
  const all = sheet.getRange(`A1:${lastCol}${last}`);
  all.format.fill = '#262626';
  all.format.font = { name: 'Arial', size: 10, color: '#FFFFFF' };
  all.format.verticalAlignment = 'center';
  all.format.rowHeight = 24;
  sheet.getRange('A2').values = [[title]];
  sheet.getRange('A2').format.font = { name: 'Arial', size: 16, bold: true, color: '#FFFFFF' };
  notes.forEach((note, i) => {
    sheet.getRange(`A${i + 3}`).values = [[note]];
    sheet.getRange(`A${i + 3}:${lastCol}${i + 3}`).format.wrapText = false;
    sheet.getRange(`A${i + 3}`).format.font = { name: 'Arial', size: 9, color: '#FFFFFF' };
    sheet.getRange(`A${i + 3}:${lastCol}${i + 3}`).format.rowHeight = 26;
  });
  sheet.getRange(`A10:${lastCol}10`).values = [headers];
  sheet.getRange(`A11:${lastCol}${last}`).values = rows;
  sheet.tables.add(`A10:${lastCol}${last}`, true, tableName);
  sheet.getRange(`A10:${lastCol}10`).format.fill = '#51463D';
  sheet.getRange(`A10:${lastCol}10`).format.font = { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' };
  sheet.getRange(`A10:${lastCol}10`).format.wrapText = true;
  sheet.getRange(`A10:${lastCol}10`).format.rowHeight = 40;
  sheet.getRange(`A11:${lastCol}${last}`).format.wrapText = true;
  sheet.getRange(`A10:${lastCol}${last}`).format.borders = { preset: 'all', style: 'thin', color: '#171717' };
  widths.forEach((width, i) => sheet.getRange(`${letters[i]}1:${letters[i]}${last}`).format.columnWidth = width);
  rows.forEach((row, i) => sheet.getRange(`A${i + 11}:${lastCol}${i + 11}`).format.fill = i % 2 ? '#514538' : '#2D2B28');
  sheet.freezePanes.unfreeze();
}

const orderStart = r => r.event ? r.event.sourceStart || r.event.start : null;
const displayDate = x => String(x || '').split('-').reverse().join('.');
const mainRows = data.rows.map(r => [r.values[0], r.values[1], r.values[2], serial(r.values[3]), serial(orderStart(r)), r.values[4], serial(r.values[5]), r.values[6], r.values[7], r.values[8]]);
layout(main, ['Звання', 'ПІБ', 'ІПН/Установа', 'Початок (таблиця)', 'Початок за наказом', 'Тривалість', 'Припинення', 'Підстава', 'Наказ', 'Додаткова інформація'], mainRows, [23, 43, 16, 17, 19, 12, 14, 76, 10, 48], 'СУХПРОД — сверка с приказами с августа 2026', [
  `Проверено ${data.files.length} приказов А7383 за 01.08.2026–${latestDate}; извлечено ${assigned} персональных назначений. Источник: РУХ_last.xlsx.`,
  `Зелёный — ${added} добавленных строк; фиолетовый — ${dateDifferences} расхождений даты начала. Даты приказа рядом с текущими, жирным. Сохранено ${data.existing.length} исходных записей.`,
  `Підстава — полный абзац; Наказ — только номер. Поля заполнены для ${matched} сопоставленных и ${added} новых записей.`,
  'До августа — исходный порядок; затем блоки и ФИО по последовательности приказов; неподтверждённые строки — в конце. Припинення = текущий Початок + Тривалість.',
  monthNote,
  `${unconfirmed} существующих строк не подтверждены доступными приказами; подробности на листе «Сверка». До августа — вне проверки.`
], 'СУХПРОД');
const end = mainRows.length + 10;
main.getRange(`D11:D${end}`).setNumberFormat('dd.mm.yyyy');
main.getRange(`E11:E${end}`).setNumberFormat('dd.mm.yyyy');
main.getRange(`E11:E${end}`).format.font.bold = true;
main.getRange(`G11:G${end}`).setNumberFormat('dd.mm.yyyy');
main.getRange(`B11:B${end}`).format.font.color = '#FFB65C';
main.getRange(`C11:C${end}`).format.font.color = '#FFB65C';
main.getRange(`F11:F${end}`).setNumberFormat('0');
main.getRange(`I11:I${end}`).setNumberFormat('0');
data.rows.forEach((r, i) => {
  const row = i + 11;
  main.getRange(`A${row}:J${row}`).format.rowHeight = r.event ? 72 : 32;
  if (r.event && !r.status.startsWith('Расхождение даты')) main.getRange(`G${row}`).formulas = [[`=D${row}+F${row}`]];
  if (r.status === 'Добавлено') main.getRange(`A${row}:J${row}`).format.fill = '#315D40';
  if (r.dateDifference) {
    main.getRange(`A${row}:J${row}`).format.fill = '#703078';
    main.getRange(`E${row}`).format.fill = '#943A87';
  }
  if (r.event?.sourceStart) {
    main.getRange(`J${row}`).values = [[`${r.values[8] || ''}${r.values[8] ? '\n' : ''}Месяц требует проверки: в №257 указан август, в таблице — сентябрь.`]];
  }
});

const priority = r => r.dateDifference ? 0 : r.status === 'Добавлено' ? 1 : r.changes ? 2 : !r.event ? 3 : 4;
const audited = data.rows.map((r, i) => ({...r, resultRow: i + 11})).filter(r => r.status !== 'Вне периода проверки').sort((a,b) => priority(a)-priority(b) || a.resultRow-b.resultRow);
const auditRows = audited.map(r => [r.status, r.values[1], r.sourceRow, r.resultRow, displayDate(r.values[3]), displayDate(orderStart(r)) || null, r.values[4], r.event?.order ?? null, r.changes || null, r.event?.note || null, r.event?.file || null, r.event?.paragraph ?? null, r.event?.original || null]);
layout(audit, ['Результат сверки', 'ПІБ', 'Строка исходной таблицы', 'Строка результата', 'Початок (таблиця)', 'Початок за наказом', 'Доби', 'Наказ', 'Изменения', 'Примечание', 'Документ приказа', 'Абзац', 'ФИО в приказе'], auditRows, [36,43,15,15,17,19,9,10,70,75,47,10,46], 'Сверка — результаты и расхождения', [
  `Сначала ${dateDifferences} расхождений даты начала (фиолетовым), затем новые строки, изменения и подтверждённые назначения.`,
  `Сопоставлено ${matched} записей; добавлено ${added}; изменены звание, длительность или прекращение в ${data.counts['Исправлено по приказу'] || 0} существующих записях.`,
  'Строки исходной таблицы относятся к сохранённой копии РУХ_last.xlsx. Строка результата — номер на листе «Сухпрод».',
  'Сначала точные совпадения ФИО, даты и части. Затем взаимно однозначные несовпадения даты в пределах 7 дней при одинаковой длительности; №257 — отдельный спорный месяц.',
  `Орфографические варианты и раскрытые инициалы отмечены. ${monthNote}`,
  'Нет источника — запись сохранена; отсутствие подтверждения не означает, что назначение ошибочно.'
], 'DryRationsAudit');
audit.getRange(`E11:E${auditRows.length + 10}`).setNumberFormat('dd.mm.yyyy');
audit.getRange(`B11:B${auditRows.length + 10}`).format.font.color = '#FFB65C';
audit.getRange(`F11:F${auditRows.length + 10}`).format.font.bold = true;
audit.getRange(`A11:M${auditRows.length + 10}`).format.rowHeight = 64;
audited.forEach((r, i) => {
  if (r.status === 'Добавлено') audit.getRange(`A${i+11}:M${i+11}`).format.fill = '#315D40';
  if (r.dateDifference) audit.getRange(`A${i+11}:M${i+11}`).format.fill = '#703078';
});

const orderRows = data.groups.map(g => [g.order, g.file, g.paragraph, serial(g.start), g.days, g.unit, g.people.length, g.declared ?? null, g.heading, g.basis, [g.order === 257 && g.start === '2026-08-05' ? 'Спорный месяц: в приказе август; соответствующие 50 строк таблицы — сентябрь.' : '', ...data.warnings.filter(w => w[0] === g.file && String(w[2]).startsWith('Число')).map(w => w[2])].filter(Boolean).join(' ') || null]);
layout(orders, ['Наказ', 'Документ', 'Абзац раздела', 'Початок по тексту', 'Доби', 'Частина', 'Людей в списке', 'Итог в приказе', 'Текст назначения', 'Підстава', 'Примечание'], orderRows, [10,47,14,17,9,13,15,15,85,85,65], 'Пункты приказов — источники сверки', [
  `${data.groups.length} групп назначений в ${new Set(data.groups.map(g => g.order)).size} приказах, включая индивидуальные назначения.`,
  `Проверены все ${data.files.length} доступных приказов начиная с 01.08.2026. Последний доступный документ — ${latestDate} №${latestOrder}.`,
  'Путь документов: Documents/Orders/А7383/2026. Номер абзаца включает пустые абзацы DOCX и служит ориентиром.',
  `Использован фактический перечень ФИО. Расхождения итоговых чисел: ${data.warnings.filter(w => String(w[2]).startsWith('Число')).map(w => '№' + w[0].match(/№\s*(\d+)/)[1] + ': ' + w[2]).join('; ') || 'нет'}.`,
  'Период перед списком применяется ко всем ФИО до соответствующего абзаца «Підстава».',
  monthNote
], 'DryRationsOrders');
orders.getRange(`D11:D${orderRows.length+10}`).setNumberFormat('dd.mm.yyyy');
orders.getRange(`A11:K${orderRows.length+10}`).format.rowHeight = 135;
main.tabColor = '#315D40';
audit.tabColor = '#51463D';
orders.tabColor = '#77604C';
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:'Сухпрод',range:'A10:I12',maxChars:1600,tableMaxCellChars:100})).ndjson);
const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(`${dir}/СУХПРОД — сверка с приказами.xlsx`);
console.log('XLSX exported');
const preview = await wb.render({sheetName:'Сверка',range:'A1:I14',scale:1,format:'png'});
await fs.writeFile(`${dir}/preview.png`, new Uint8Array(await preview.arrayBuffer()));
console.log('Preview exported');