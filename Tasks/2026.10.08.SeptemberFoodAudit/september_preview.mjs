import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from 'file:///C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';

const directory = 'C:/WORK/Windows/Строевые приказы/Tasks/2026.10.08.SeptemberFoodAudit/Results';
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(`${directory}/Продовольче — сверка сентября 2026.xlsm`));
const preview = await workbook.render({ sheetName: 'Котел', range: 'A436:H441', scale: 1.5, format: 'png' });
await fs.writeFile(`${directory}/preview.png`, new Uint8Array(await preview.arrayBuffer()));
console.log('Preview saved');