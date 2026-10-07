import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {createRequire} from 'node:module';
import {execFileSync} from 'node:child_process';

const here=path.dirname(fileURLToPath(import.meta.url));
const args=process.argv.slice(2);
const option=(name,fallback)=>args.includes(name)?args[args.indexOf(name)+1]:fallback;
const sourcesDir=option('--sources-dir',path.dirname(here));
const movement=option('--movement');
const month=option('--month','2026-09');
let output=option('--output',path.join(sourcesDir,'Результат',`СЗЧ_${month}_с_итоговым_списком.xlsx`));
if(!movement)throw new Error('Укажите --movement с путём к РУХ_last.xlsx');
const runtime=path.join(process.env.USERPROFILE,'.cache','codex-runtimes','codex-primary-runtime','dependencies');
const python=process.env.SZCH_PYTHON||path.join(runtime,'python','python.exe');
const nodeModules=process.env.SZCH_NODE_MODULES||path.join(runtime,'node','node_modules');
const require=createRequire(path.join(nodeModules,'package.json'));
const {Workbook,SpreadsheetFile}=await import(pathToFileURL(require.resolve('@oai/artifact-tool')).href);
const data=JSON.parse(execFileSync(python,[path.join(here,'analyze.py'),'--sources-dir',sourcesDir,'--movement',movement,'--month',month,'--full'],{encoding:'utf8',env:{...process.env,PYTHONIOENCODING:'utf-8'},maxBuffer:20000000}));
const people=data.people,found=people.filter(p=>p.matches.length),final=people.filter(p=>p.final);
if(!people.length)throw new Error('За выбранный месяц нет зачислений');
const wb=Workbook.create();
const stages=wb.worksheets.add('Этапы фильтрации');
const audit=wb.worksheets.add('Сверка ФИО');
const history=wb.worksheets.add('История ТП');
const colors={background:'#262626',row:'#2D2D2D',alternate:'#383838',header:'#365D89',text:'#F2F2F2',pink:'#E9B8C0',pinkText:'#541C26',green:'#A6D65C'};
const date=x=>x?new Date(x.endsWith('Z')?x:x+'Z'):null;
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const end=people.length+9;
function base(sheet,last,bottom,title){
 sheet.showGridLines=false;sheet.freezePanes.unfreeze();sheet.tabColor=colors.header;
 sheet.getRange(`A1:${last}${bottom}`).format={fill:colors.background,font:{name:'Arial',size:10,color:colors.text},rowHeight:25,verticalAlignment:'center'};
 sheet.getRange('A2').values=[[title]];sheet.getRange('A2').format.font={size:14,bold:true};
}
function table(sheet,address,name){
 const t=sheet.tables.add(address,true,name);t.style='TableStyleLight1';t.showFilterButton=true;
 t.getHeaderRowRange().format={fill:colors.header,font:{bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:44};return t;
}
function body(sheet,start,last,count){
 for(let i=0;i<count;i++)sheet.getRange(`${start}${i+10}:${last}${i+10}`).format.fill=(i+10)%2===0?colors.alternate:colors.row;
}
const source=s=>`${s.source.startsWith('00000.')?'БРЕЗ':s.source.startsWith('0000.')?'Суд':'Рапорты'}!${s.row}`;
base(audit,'R',end+20,'Сверка всех зачислений за '+month);
audit.getRange('A3').values=[['Суд = 0000. СЗЧ по ухвалі суду 05_10_2026.xlsx / Лист1']];
audit.getRange('A4').values=[['БРЕЗ = 00000. СЗЧ_через_БРЕЗ_новий_алгоритм_на 05.10.2026.xlsx / Лист1']];
audit.getRange('A5').values=[['Рапорты = Поіменний список рапортів СЗЧ (1).xlsm / Аркуш1']];
audit.getRange('A6').values=[['ФИО нормализуется. При отсутствии точного совпадения проверяется ИПН; ближайшее написание сохраняется для проверки.']];
audit.getRange('A7').values=[['Разные фамилии или ИПН не объединяются автоматически. Сходство текста не является вероятностью совпадения.']];
const details=people.map(p=>{
 const matched=!!p.matches.length,s=matched?p.matches[0]:p.candidates[0];
 const conflict=p.matches.some(x=>x.ipn&&p.ipn&&x.ipn!==p.ipn);
 const result=p.final?'Включён':!matched?'Не найден в трёх списках':!p.selected?'Не определена дата ТП':'Прибыл вне выбранного месяца';
 const note=conflict?'Расхождение ИПН при точном ФИО':p.match_type==='ИПН'?'Вариант имени подтверждён ИПН':matched?'':s?.ipn&&p.ipn&&s.ipn!==p.ipn?'Ближайшее ФИО имеет другой ИПН':'Нет достаточного подтверждения';
 return [p.name,p.ipn,date(p.enroll),p.row,matched?'Да':'Нет',p.match_type,s?.name||'',s?.ipn||'',matched&&p.match_type==='ФИО'?1:p.candidates.find(x=>x.name===s?.name)?.score??null,(matched?p.matches:s?[s]:[]).map(source).join('\n'),date(p.latest?.arrival),date(p.selected?.arrival),p.selected?.row??null,p.latest&&p.selected&&p.latest.row!==p.selected.row?'Да':'Нет',result,note];
});
audit.getRange(`A9:P${end}`).values=[['ПІБ в РУХ','ІПН в РУХ','Зарахування','Строка Список','В списках СЗЧ','Основание совпадения','ФИО источника / ближайшее','ИПН источника','Сходство текста','Источник и строка','Последняя дата ТП','Фактическое прибытие','Строка ТП','Учтена смена статуса','Результат','Примечание'],...details];
table(audit,`A9:P${end}`,'NameAudit');body(audit,'A','P',people.length);
[48,15,15,13,14,19,48,16,14,24,16,20,12,18,30,49].forEach((w,i)=>audit.getRange(`${col(i)}1:${col(i)}${end}`).format.columnWidth=w);
for(const c of ['C','K','L'])audit.getRange(`${c}10:${c}${end}`).setNumberFormat('dd.mm.yyyy');
audit.getRange(`I10:I${end}`).setNumberFormat('0.0%');audit.getRange(`J10:J${end}`).format.wrapText=true;
for(let i=0;i<people.length;i++){
 if(people[i].final)audit.getRange(`O${i+10}`).format={fill:'#304329',font:{color:colors.green}};
 if(details[i][15].startsWith('Расхождение'))audit.getRange(`P${i+10}`).format={fill:'#493727',font:{color:'#FFA45B'}};
}
base(stages,'Q',end+28,'СЗЧ — фильтрация за '+month);
stages.getRange('A3').values=[['Зачисление за выбранный месяц. Совпадение хотя бы с одним из трёх списков.']];
stages.getRange('A4').values=[['Первые три блока — полный список в одном порядке. Розовый фон — исключён к этому этапу.']];
const groups=[people,people,people,final];
const titles=['1. Зачислены за месяц','2. Найдены в списках СЗЧ','3. Прибыли за месяц — итог','4. Чистый итоговый список'];
for(let g=0;g<4;g++){
 const c=g*4,a=col(c),b=col(c+1),z=col(c+2),rows=groups[g],last=rows.length+9;
 stages.getRange(`${a}6:${z}6`).format={fill:colors.alternate,font:{bold:true},rowHeight:30};stages.getRange(`${a}6`).values=[[titles[g]]];
 stages.getRange(`${a}7`).formulas=[[g===0?`=COUNTA(A10:A${end})`:g===1?`=COUNTIFS('Сверка ФИО'!$E$10:$E$${end},"Да")`:g===2?`=COUNTIFS('Сверка ФИО'!$O$10:$O$${end},"Включён")`:`=COUNTA(M10:M${Math.max(10,last)})`]];
 stages.getRange(`${a}7`).format.font.color=colors.green;stages.getRange(`${b}7`).values=[[g===3?'человек':'прошли этап']];
 if(g<3)stages.getRange(`${a}8`).values=[[`Исключено: ${g===0?0:g===1?people.length-found.length:people.length-final.length}${g===2?', из них '+(found.length-final.length)+' на этапе ТП':''}`]];
 stages.getRange(`${a}9:${z}9`).values=[['ПІБ','ІПН',g<2?'Зарахування':'Фактическое прибытие']];
 if(rows.length){
  stages.getRange(`${a}10:${z}${last}`).values=rows.map(p=>[p.name,p.ipn,g<2?date(p.enroll):date(p.selected?.arrival)]);
  body(stages,a,z,rows.length);
  for(let i=0;i<rows.length;i++)if((g===1&&!rows[i].matches.length)||(g===2&&!rows[i].final))stages.getRange(`${a}${i+10}:${z}${i+10}`).format={fill:colors.pink,font:{color:colors.pinkText}};
  stages.getRange(`${b}10:${b}${last}`).setNumberFormat('@');stages.getRange(`${z}10:${z}${last}`).setNumberFormat('dd.mm.yyyy');
 }
 table(stages,`${a}9:${z}${Math.max(10,last)}`,g===3?'FinalPeople':`Stage${g+1}`);
 stages.getRange(`${a}1:${a}${end}`).format.columnWidth=30;stages.getRange(`${b}1:${b}${end}`).format.columnWidth=15;stages.getRange(`${z}1:${z}${end}`).format.columnWidth=14;
 stages.getRange(`${b}9:${z}${Math.max(10,last)}`).format.horizontalAlignment='center';
 if(g<3)stages.getRange(`${col(c+3)}1:${col(c+3)}${end}`).format.columnWidth=3;
}
const events=found.flatMap(p=>p.events.map(t=>[p.name,t.row,date(t.arrival),date(t.departure),t.row===p.selected?.row?'Выбрана':t.row===p.latest?.row?'Последняя запись':'Предыдущая запись',t.note||'',t.name,t.basis||'',t.depart_basis||'',p.final?'Включён':'Прибыл вне выбранного месяца']));
base(history,'L',events.length+28,'Выбор фактической даты прибытия');
history.getRange('A3').values=[['Источник: РУХ_last.xlsx / Прибулі / таблица ТП; номера строк исходного листа.']];
history.getRange('A4').values=[['Последнее прибытие не позже зачисления; связанная смена статуса отсылает к предыдущей записи.']];
history.getRange('A5').values=[['Связь: предыдущее выбытие равно следующему прибытию; примечание содержит «Зміна статуса».']];
history.getRange(`A8:J${events.length+8}`).values=[['ПІБ в Списке','Строка ТП','Прибуття','Вибуття','Выбор даты','Додаткова інформація','ПІБ в ТП','Прибуття.Підстава','Вибуття.Підстава','Результат'],...events];
table(history,`A8:J${Math.max(9,events.length+8)}`,'ArrivalHistory');
[48,12,15,15,23,80,48,85,85,28].forEach((w,i)=>history.getRange(`${col(i)}1:${col(i)}${events.length+8}`).format.columnWidth=w);
history.getRange(`C9:D${Math.max(9,events.length+8)}`).setNumberFormat('dd.mm.yyyy');history.getRange(`F9:I${Math.max(9,events.length+8)}`).format.wrapText=true;
for(let i=0;i<events.length;i++){
 history.getRange(`A${i+9}:J${i+9}`).format.fill=(i+9)%2===0?colors.alternate:colors.row;
 history.getRange(`A${i+9}:J${i+9}`).format.rowHeight=Math.min(220,Math.max(48,Math.ceil(Math.max(String(events[i][5]).length/78,String(events[i][7]).length/84,String(events[i][8]).length/84))*15));
 if(events[i][4]==='Выбрана')history.getRange(`E${i+9}`).format={fill:'#304329',font:{color:colors.green}};
}
wb.recalculate();
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!',options:{useRegex:true,maxResults:10},summary:'Проверка формул'})).ndjson);
await fs.mkdir(path.dirname(output),{recursive:true});
const temporary=path.join(path.dirname(output),'.szch-preview.png');
const preview=await wb.render({sheetName:stages.name,range:'I6:O17',scale:1,format:'png'});await fs.writeFile(temporary,new Uint8Array(await preview.arrayBuffer()));
try{await (await SpreadsheetFile.exportXlsx(wb)).save(output);}catch(error){if(error.code!=='EBUSY')throw error;output=output.replace(/\.xlsx$/i,`_${Date.now()}.xlsx`);await (await SpreadsheetFile.exportXlsx(wb)).save(output);console.log('Книга открыта; сохранено:',output);}
console.log(JSON.stringify({month,counts:[people.length,found.length,final.length],sources:data.source_counts,output,preview:temporary}));