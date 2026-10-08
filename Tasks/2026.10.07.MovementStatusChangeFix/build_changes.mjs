import fs from 'node:fs/promises';
import {spawnSync} from 'node:child_process';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';

const folder='C:/WORK/Windows/Строевые приказы/Tasks/2026.10.07.MovementStatusChangeFix';
const existingVersions=(await fs.readdir(folder)).map(name=>name.match(/^Анализ смен статусов v(\d+)\.xlsx$/)).filter(Boolean).map(match=>Number(match[1]));
const nextOutputPath=`${folder}/Анализ смен статусов v${Math.max(0,...existingVersions)+1}.xlsx`;
const result=spawnSync('C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',[`${folder}/decide_changes.py`],{encoding:'utf8',maxBuffer:40*1024*1024});
if(result.status!==0)throw new Error(result.stderr);
const data=JSON.parse(result.stdout);
const singlePoint=process.argv.includes('--single-point');
const splitTypes=process.argv.includes('--split-types');
const changesOnly=process.argv.includes('--changes-only');
const compact=process.argv.includes('--compact');
const fourRows=process.argv.includes('--four-rows');
const version3=process.argv.includes('--v3');
const dischargeRule=process.argv.includes('--discharge-rule');
const originalCount=data.decisions.length;
function personPoints(decision){
 const points=new Set();
 for(const match of decision.case.matches){
  const headings=[...match.context,match.text].map(text=>text.match(/^(\d+(?:\.\d+)*)\.\s/)).filter(Boolean);
  if(headings.length)points.add(headings.at(-1)[1]);
  else points.add(`без номера:${match.path}:${match.paragraph}`);
 }
 return points;
}
if(singlePoint)data.decisions=data.decisions.filter(decision=>{
 const points=personPoints(decision);
 return points.size===1&&![...points][0].startsWith('без номера:');
});
const unchangedPairCount=changesOnly?data.decisions.filter(decision=>decision.sides.every(side=>side.verdict==='Без изменений')&&!decision.additions.length).length:0;
if(changesOnly)data.decisions=data.decisions.filter(decision=>decision.sides.some(side=>side.verdict!=='Без изменений')||decision.additions.length);
const peopleCount=new Set(data.decisions.map(d=>d.pair[1]||d.pair[0])).size;
const orderCount=new Set(data.decisions.map(d=>`${d.pair[6].slice(0,4)}:${d.pair[4]}`)).size;
const wb=Workbook.create();
const main=wb.worksheets.add('Сверка событий');
const serial=value=>value?Date.parse(value.slice(0,10)+'T00:00:00Z')/86400000+25569:null;
const sources=[];
const sourceMap=new Map();
function evidence(match,person,order){
 const key=match.path+'|'+match.paragraph;
 if(sourceMap.has(key))return sourceMap.get(key);
 const id='П'+String(sources.length+1).padStart(4,'0');
 const section=match.context.map(v=>v.match(/^(\d+(?:\.\d+)*)\.\s/)).filter(Boolean).map(v=>v[1]).join('; ');
 sources.push([id,person,Number(order),serial(match.order_date),section,match.paragraph,match.source,match.basis,match.path,match.identity]);
 sourceMap.set(key,id);
 return id;
}
function priority(decision){
 if(decision.sides.some(s=>s.verdict.startsWith('Исправить')))return 0;
 if(decision.sides.some(s=>s.verdict.startsWith('Предлагается'))||decision.additions.length)return 1;
 if(decision.sides.some(s=>s.verdict!=='Без изменений'))return 2;
 return 3;
}
const ordered=[...data.decisions].sort((a,b)=>priority(a)-priority(b)||a.pair[0].localeCompare(b.pair[0],'uk')||a.id-b.id);
const rows=[];
for(const decision of ordered){
 const p=decision.pair;
 for(const [index,side] of decision.sides.entries()){
  let matches=side.chosen?[side.chosen.match]:decision.case.matches.filter(m=>!m.source.match(/Виплатити матеріальну допомогу|Виплачувати грошову допомогу/i));
  const ids=[...new Set(matches.map(m=>evidence(m,p[0],p[4])))];
  const chosen=side.chosen?.match;
  const date=chosen?.order_date||decision.case.files.find(v=>v.verified)?.date||decision.case.files[0]?.date;
  const sections=[...new Set(matches.flatMap(m=>m.context.map(v=>v.match(/^(\d+(?:\.\d+)*)\.\s/)).filter(Boolean).map(v=>v[1])))].join('; ');
  const paths=[...new Set(matches.map(m=>m.path))];
  rows.push([decision.id,p[0],p[1],side.source_row,index===0?'Закрытие: Прибуття':'Открытие: Вибуття',side.original_status,serial(side.original_date),side.corrected_status||'Не определено',serial(side.corrected_date),null,side.verdict,side.reason,p[index===0?10:11],Number(p[4]),serial(date),sections,ids.join('; '),paths.length?paths.join('\n'):decision.case.files.map(v=>v.path).join('\n')]);
 }
 for(const addition of decision.additions){
  const match=addition.match;
  const id=evidence(match,p[0],p[4]);
  const section=match.context.map(v=>v.match(/^(\d+(?:\.\d+)*)\.\s/)).filter(Boolean).map(v=>v[1]).join('; ');
  rows.push([decision.id,p[0],p[1],null,'Добавить интервал',null,null,addition.status,serial(addition.from),serial(addition.to),'Добавить промежуточное событие',addition.reason,null,Number(p[4]),serial(match.order_date),section,id,match.path]);
 }
}
if(fourRows){
 const reader=`import sys,json,openpyxl,datetime
sys.stdout.reconfigure(encoding='utf-8')
w=openpyxl.load_workbook(sys.argv[1],read_only=True,data_only=True)
s=w['Відсутні']
print(json.dumps({str(i):list(r[:17]) for i,r in enumerate(s.iter_rows(values_only=True),1) if i>=7},ensure_ascii=False,default=lambda v:v.isoformat() if isinstance(v,(datetime.date,datetime.datetime)) else str(v)))`;
 const read=spawnSync('C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',['-c',reader,`${folder}/РУХ_source_snapshot.xlsx`],{encoding:'utf8',maxBuffer:40*1024*1024});
 if(read.status!==0)throw new Error(read.stderr);
 const original=JSON.parse(read.stdout);
 const dateSame=(a,b)=>String(a||'').slice(0,10)===String(b||'').slice(0,10);
 for(const d of data.decisions){
  const p=d.pair,a=original[p[8]],b=original[p[9]];
  if(!a||!b||String(a[3]||'')!==String(p[1]||'')||String(b[3]||'')!==String(p[1]||'')||(!p[1]&&(a[2]!==p[0]||b[2]!==p[0]))||a[5]!==p[2]||b[5]!==p[3]||!dateSame(a[14],p[5])||!dateSame(b[10],p[6])||String(a[16])!==String(p[4])||String(b[8])!==String(p[4]))throw new Error(`Исходная копия не соответствует паре ${d.id}`);
 }
 const grouped=[...ordered].sort((a,b)=>a.pair.slice(2,4).join(' → ').toLocaleLowerCase('uk').localeCompare(b.pair.slice(2,4).join(' → ').toLocaleLowerCase('uk'),'uk')||priority(a)-priority(b)||a.pair[0].localeCompare(b.pair[0],'uk')||a.id-b.id);
 const view=[];
 const displayDate=value=>value?value.slice(0,10).split('-').reverse().join('.'):'';
 function detailedAnalysis(d,side){
  const p=d.pair;
  const input=`В таблице: «${p[2]}», Прибуття ${displayDate(p[5])}; затем «${p[3]}», Вибуття ${displayDate(p[6])}. На переход указан приказ №${p[4]}.`;
  const specific={
   27:'Проверить ІПН и полное отчество в личной карточке и первичном документе: Євгенович и Євгенійович могут обозначать разных людей. До подтверждения личности нельзя переносить даты из этого пункта в запись таблицы.',
   67:"Сопоставить ІПН, полное имя и подразделение в личной карточке и основании приказа. Установить, является ли ВОРОБ’Я падежной формой фамилии ВОРОБЕЙ или относится к другому человеку.",
   89:'Проверить полный пункт 70 и соседние заголовки в оригинале, затем рапорт и документ о направлении: какое действие предписано, какой статус прекращается и какой начинается? Срок и место сами по себе не подтверждают смену статуса.',
   191:'Пункт 14 приказа №369 от 19.12.2025 предписывает отпуск для лечения с 20.12.2025 по 17.02.2026. Проверяемая пара — окончание отпуска 14.09.2025 и начало стационарного лечения 15.09.2025. Расходятся и период, и направление перехода. Проверить номер приказа в обеих строках ТВ; найти сентябрьский приказ о госпитализации и прекращении отпуска, сверить дату поступления по медицинскому документу. По декабрьскому пункту нельзя исправлять сентябрьские даты или объявлять ошибку в самом приказе.',
   210:'Проверить ІПН и исходное написание фамилии в личной карточке и рапорте: ЛОМАКО и ЛОМАЦІ нельзя автоматически считать одним человеком. Только после подтверждения сопоставить даты и статус.',
   272:'Открыть оригинал и проверить подписанный заголовок, регистрационный номер и дату. Имя файла указывает №162 от 03.06.2026, заголовок — №161 от 02.06.2026. Установить, ошибочно ли имя файла, ссылка в таблице или выбран документ; до этого даты из пункта не переносить.',
   334:'Сверить год выписки с медицинской выпиской: в пункте 2035, в основании 2025. Проверить вид отпуска по решению ВЛК, рапорту и отпускному билету: заголовок говорит об отпуске для лечения, персональный текст — о ежегодном. Установить правильные год и вид отпуска до исправления таблицы и текста приказа.',
   440:'Проверить предыдущий приказ и медицинскую выписку: перед новым событием человек находился в стационаре или уже в отпуске для лечения? В таблице и персональной формулировке прежний статус разный. Выписка в основании не доказывает, что непосредственно предыдущее событие было стационарным лечением.',
   265:'Уточнить цель и характер направления в медицинскую роту по рапорту, отпускному билету и медицинскому документу: стационар, амбулаторное обследование или отпуск. Само название медицинской роты не определяет статус; проверить дату начала подтверждённого события.'
  };
  let check=specific[d.id];
  if(!check&&side.verdict==='Уточнить тип медицинского события')check='Проверить по медицинской выписке или справке ВЛК, было ли круглосуточное стационарное лечение либо амбулаторный осмотр. Название учреждения не устанавливает вид события. Сверить даты поступления, выписки и приказ направления.';
  if(!check&&side.verdict==='Проверить противоречие')check='Сопоставить перечисленные даты с персональными пунктами каждой версии приказа и первичным основанием. Установить, относятся ли они к одному переходу или к разным событиям; выбрать дату только после устранения противоречия.';
  if(!check&&/Проверить дату/.test(side.verdict))check='Сверить номер и год приказа, ПІБ/ІПН, персональный пункт и медицинский документ или рапорт. Существенная разница дат может означать неверную ссылку либо иной период, а не опечатку. Сначала найти документ именно на исходный переход.';
  if(!check&&side.verdict==='Требуется проверка')check=`В пункте не удалось однозначно подтвердить ${side.direction==='Прибытие'?'закрытие':'открытие'} статуса «${side.original_status}» на ${displayDate(side.original_date)}. Проверить формулировку действия, прежний и новый статусы, фактическую дату; сопоставить предыдущий/следующий приказ и первичное основание. Не считать дату приказа или продовольственную дату доказательством фактического события.`;
  if(!check&&side.verdict==='Предлагается учётная дата')check='Дата закрытия прежнего статуса выведена из начала нового, а не из отдельной команды о возвращении. Проверить, не указаны ли отдельно выписка, прибытие или прекращение прежнего статуса. Если отпуск начинается в последний день лечения, сначала исправить период по правилу выписки; не выравнивать даты автоматически.';
  if(!check&&side.verdict==='Предлагается дата приказа')check='Фактическая дата действия прямо не найдена. Проверить рапорт и первичный документ, а также общий заголовок группы. Дата приказа здесь только предположение; продовольственная дата не подтверждает фактическую.';
  const candidates=[...new Set(side.candidates.map(c=>`«${c.status}» — ${displayDate(c.date)} (${c.kind})`))];
  return [input,candidates.length?`Найденные варианты: ${candidates.join('; ')}.`:'',check,check&&/Требуется|Уточнить|Проверить/.test(side.verdict)?'«Не определено» и пустая дата означают отсутствие подтверждённого исправления; исходную запись не удалять и не заменять пустыми значениями.':''].filter(Boolean).join('\n\n');
 }
 for(const d of grouped){
  const p=d.pair;
  const matches=[...new Map(d.case.matches.map(m=>[m.path+'|'+m.paragraph,m])).values()];
  let text=matches.map(m=>`${m.source}${m.basis?'\n\n'+m.basis:''}\n\nИсточник: ${version3?m.path.split(/[\\/]/).at(-1):m.path}${m.identity.startsWith('Возможное')?'\n'+m.identity:''}`).join('\n\n');
  let correctedText=text,overlapDate=null;
  if(dischargeRule&&p[2]==='Стаціонарне лікування'&&p[3].toLowerCase()==='відпустка для лікування'){
   const months=['січня','лютого','березня','квітня','травня','червня','липня','серпня','вересня','жовтня','листопада','грудня'];
   const treatment=text.match(/по\s+(\d{1,2})\s+([а-яіїєґ]+)\s+(\d{4})\s+року\s+перебував\s+на\s+лікуванні/i);
   const leave=text.slice(text.search(/вваж[аи]ти\s+таким/i)).match(/з\s+(\d{1,2})\s+([а-яіїєґ]+)\s+по\s+(\d{1,2})\s+([а-яіїєґ]+)\s+(\d{4})\s+року/i);
   if(treatment&&leave&&Number(treatment[1])===Number(leave[1])&&treatment[2]===leave[2]&&treatment[3]===leave[5]&&months.includes(leave[2])){
    const start=new Date(Date.UTC(Number(leave[5]),months.indexOf(leave[2]),Number(leave[1])+1));
    const end=new Date(Date.UTC(Number(leave[5]),months.indexOf(leave[4]),Number(leave[3])+1));
    overlapDate=start.toISOString().slice(0,10);
    const replacement=`з ${String(start.getUTCDate()).padStart(2,'0')} ${months[start.getUTCMonth()]} по ${String(end.getUTCDate()).padStart(2,'0')} ${months[end.getUTCMonth()]} ${end.getUTCFullYear()} року`;
    correctedText=text.replaceAll(leave[0],replacement);
   }
  }
  for(const corrected of [false,true])for(let sideIndex=0;sideIndex<2;sideIndex++){
   const side=d.sides[sideIndex],raw=original[p[8+sideIndex]],values=[...raw];
   if(corrected){
    values[5]=side.corrected_status||'Не определено';
    values[sideIndex===0?14:10]=side.corrected_date;
    if(overlapDate){values[5]=raw[5];values[sideIndex===0?14:10]=overlapDate;}
   }
   const extra=d.additions.map(a=>`Добавить промежуточное событие «${a.status}»: ${displayDate(a.from)} — ${displayDate(a.to)}. ${a.reason}`).join('\n');
   const operation=[corrected?'Исправленная строка':'Исходная строка',side.direction,overlapDate?'Исправление по правилу выписки':side.verdict,overlapDate?`В приказе начало отпуска совпадает с последним днём лечения. Предлагается начало ${displayDate(overlapDate)} и перенос конца отпуска на 1 день с сохранением длительности. Прибуття лечения и Вибуття отпуска: ${displayDate(overlapDate)}. Остальные даты события сохранены.`:side.reason,detailedAnalysis(d,side),extra].filter(Boolean).join('\n\n');
   const orderStatus=overlapDate?'Ошибка в приказе: отпуск начинается в последний день лечения':d.sides.some(s=>/Требуется|Уточнить|Проверить/.test(s.verdict))?'Требует проверки':null;
   view.push([p.slice(2,4).join(' → '),values[2],values[3],values[5],values[6],values[8],serial(values[9]),serial(values[10]),values[11],serial(values[14]),serial(values[15]),values[16],corrected?correctedText:text,operation,d.id,side.source_row,corrected?'Исправленная':'Исходная',orderStatus]);
  }
 }
 const bottom=10+view.length;
 setup(main,'Смена статуса — исходная и исправленная пары','R',bottom,version3?[14,18,12,13,20,11,11,11,11,11,11,11,105,65,10,14,18,44]:[36,35,16,28,32,15,18,18,20,18,18,15,105,65,10,14,18,44]);
 if(dischargeRule)main.getRange(`M1:M${bottom}`).format.columnWidth=75;
 main.getRange('A3').values=[[`${data.decisions.length} пар. В каждом блоке: 2 исходные строки, затем 2 исправленные строки на зелёном фоне.`]];
 main.getRange('A4').values=[[version3?'Ярко-зелёные ячейки с жирным текстом — изменённые значения. Продовольственные даты сохранены без пересчёта.':'Порядок колонок повторяет целевую таблицу. Column1 — исходный Вибуття.Термін. Продовольственные даты сохранены без пересчёта.']];
 main.getRange('A5').values=[['Исходные полные строки взяты из сохранённой копии РУХ; люди, статусы, даты переходов и номера приказов сверены с исходной выборкой.']];
 main.getRange('A10:R10').values=[['Тип пары событий','ПІБ','ІПН','Подія','Куди','Вибуття.Наказ','Вибуття.Продовольче','Вибуття',version3?'Кол-во дней / план прибытия':'Column1','Прибуття','Прибуття.Продовольче','Прибуття.Наказ','Текст','Операция','Пара','Строка ТВ','Версия','Статус пункта приказа']];
 main.getRange(`A11:R${bottom}`).values=view;
 main.tables.add(`A10:R${bottom}`,true,'FourRowReview').style='TableStyleMedium15';
 main.getRange(`R10:R${bottom}`).format.horizontalAlignment='center';
 main.getRange(`R10:R${bottom}`).format.wrapText=true;
 main.getRange('R10').format.fill='#50463D';
 main.getRange('R10').format.font.bold=true;
 main.getRange(`A10:R${bottom}`).format.horizontalAlignment='center';
 main.getRange(`A10:R${bottom}`).format.verticalAlignment='center';
 main.getRange(`A10:R${bottom}`).format.wrapText=true;
 main.getRange('A10:R10').format.fill='#50463D';
 main.getRange('A10:R10').format.rowHeight=55;
 main.getRange('A10:R10').format.font.bold=true;
 main.getRange(`A11:R${bottom}`).format.fill='#302D29';
 main.getRange(`B11:B${bottom}`).format.font.color='#E5AE72';
 main.getRange(`C11:C${bottom}`).setNumberFormat('@');
 for(const column of ['G','H','J','K'])main.getRange(`${column}11:${column}${bottom}`).setNumberFormat('dd.mm.yyyy');
 for(let i=0;i<view.length;i++){
  const row=main.getRange(`A${i+11}:R${i+11}`);
  if(i%4>=2)row.format.fill='#34513C';
  if(version3&&i%4>=2){
   for(let column=0;column<12;column++)if(view[i][column]!==view[i-2][column]){
    const cell=main.getRangeByIndexes(i+10,column,1,1);
    cell.format.fill='#3B7D23';
    cell.format.font.bold=true;
    cell.format.font.color='#FFFFFF';
   }
  }
  if(i%4===0)row.format.borders={top:{style:'medium',color:'#A89077'}};
  const lines=String(view[i][12]).split('\n').reduce((n,t)=>n+Math.max(1,Math.ceil(t.length/(dischargeRule?68:96))),0);
  const operationLines=String(view[i][13]).split('\n').reduce((n,t)=>n+Math.max(1,Math.ceil(t.length/58)),0);
  row.format.rowHeight=Math.min(409,Math.max(110,Math.max(lines,operationLines)*14+20));
 }
 wb.recalculate();
 const preview=await wb.render({sheetName:'Сверка событий',range:'A1:N14',scale:1,format:'png'});
 await fs.writeFile(`${folder}/${version3?'v3':'v2'}-four-rows-preview.png`,new Uint8Array(await preview.arrayBuffer()));
 const outputPath=nextOutputPath;
 await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
 if(dischargeRule){
  const rich=spawnSync('C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',[`${folder}/highlight_text.py`,outputPath],{encoding:'utf8'});
  if(rich.status!==0)throw new Error(rich.stderr);
  console.log(rich.stdout);
 }
 console.log(JSON.stringify({pairs:data.decisions.length,rows:view.length,path:outputPath}));
 process.exit(0);
}
if(compact){
 const dateText=value=>value?value.slice(0,10).split('-').reverse().join('.'):'';
 const normalize=value=>value.trim().toLocaleLowerCase('uk');
 const labels=new Map();
 const decisions=new Map(data.decisions.map(d=>[d.id,d]));
 for(const d of data.decisions){
  const key=d.pair.slice(2,4).map(normalize).join(' → ');
  if(!labels.has(key))labels.set(key,d.pair.slice(2,4).join(' → '));
 }
 const evidenceById=new Map(sources.map(s=>[s[0],s]));
 const sorted=[...rows].sort((a,b)=>{
  const da=decisions.get(a[0]),db=decisions.get(b[0]);
  const ta=da.pair.slice(2,4).map(normalize).join(' → '),tb=db.pair.slice(2,4).map(normalize).join(' → ');
  return ta.localeCompare(tb,'uk')||priority(da)-priority(db)||a[1].localeCompare(b[1],'uk')||a[0]-b[0];
 });
 const compactRows=sorted.map(row=>{
  const d=decisions.get(row[0]);
  const type=labels.get(d.pair.slice(2,4).map(normalize).join(' → '));
  const fragments=row[16].split(';').map(id=>evidenceById.get(id.trim())).filter(Boolean);
  const text=[...new Set(fragments.map(s=>[s[6],s[7]?`Підстава: ${s[7].replace(/^Підстава:\s*/i,'')}`:'',`Источник: ${s[8]}`,s[9].startsWith('Возможное')?s[9]:''].filter(Boolean).join('\n\n')))].join('\n\n');
  const operation=[row[4],row[10],row[11],row[3]?`Строка ТВ: ${row[3]}`:'',row[9]?`Конец добавляемого интервала: ${dateText(new Date((row[9]-25569)*86400000).toISOString())}`:''].filter(Boolean).join('\n');
  const orderDate=row[14]?dateText(new Date((row[14]-25569)*86400000).toISOString()):'';
  return [type,row[1],row[2],row[5],row[6],row[7],row[8],row[12],`№ ${row[13]}${orderDate?` от ${orderDate}`:''}${row[15]?`\nПункт: ${row[15]}`:''}`,text||row[17],operation];
 });
 const bottom=10+compactRows.length;
 setup(main,'Смена статуса — исходные и исправленные события','K',bottom,[36,34,15,28,17,28,17,32,24,105,65]);
 main.getRange('A3').values=[[`${data.decisions.length} пар событий, ${peopleCount} человек. Один пункт о человеке в приказе; полностью неизменённые пары исключены.`]];
 main.getRange('A4').values=[['Строки сгруппированы по типу исходного перехода. Внутри пары сохранены закрытие прежнего и открытие нового статуса.']];
 main.getRange('A5').values=[['«Предлагается учётная дата» — предложение; «Исправить» — исправление по приказу. Обоснования и пункты приказов включены в таблицу.']];
 main.getRange('A10:K10').values=[['Тип пары событий','ПІБ','ІПН','Подія исходная','Дата исходная','Подія исправленная','Дата исправленная','Куди исходное','Приказ и пункт','Текст','Операция']];
 main.getRange(`A11:K${bottom}`).values=compactRows;
 main.tables.add(`A10:K${bottom}`,true,'CompactMovementReview').style='TableStyleMedium15';
 main.getRange(`A10:K${bottom}`).format.horizontalAlignment='center';
 main.getRange(`A10:K${bottom}`).format.verticalAlignment='center';
 main.getRange(`A10:K${bottom}`).format.wrapText=true;
 main.getRange('A10:K10').format.fill='#50463D';
 main.getRange('A10:K10').format.font.bold=true;
 main.getRange('A10:K10').format.rowHeight=48;
 main.getRange(`A11:K${bottom}`).format.fill='#302D29';
 main.getRange(`B11:B${bottom}`).format.font.color='#E5AE72';
 main.getRange(`C11:C${bottom}`).setNumberFormat('@');
 for(const column of ['E','G'])main.getRange(`${column}11:${column}${bottom}`).setNumberFormat('dd.mm.yyyy');
 for(let i=0;i<compactRows.length;i++){
  const lines=(value,width)=>String(value||'').split('\n').reduce((sum,line)=>sum+Math.max(1,Math.ceil(line.length/width)),0);
  main.getRange(`A${i+11}:K${i+11}`).format.rowHeight=Math.min(409,Math.max(100,Math.max(lines(compactRows[i][9],95),lines(compactRows[i][10],58))*14+20));
  if(i===0||compactRows[i][0]!==compactRows[i-1][0])main.getRange(`A${i+11}:K${i+11}`).format.borders={top:{style:'medium',color:'#A89077'}};
 }
 wb.recalculate();
 console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!',options:{useRegex:true,maxResults:10},maxChars:500})).ndjson);
 const preview=await wb.render({sheetName:'Сверка событий',range:'A1:K13',scale:1,format:'png'});
 await fs.writeFile(`${folder}/v2-compact-preview.png`,new Uint8Array(await preview.arrayBuffer()));
 const outputPath=nextOutputPath;
 await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
 console.log(JSON.stringify({pairs:data.decisions.length,rows:compactRows.length,types:labels.size,path:outputPath}));
 process.exit(0);
}
function setup(sheet,title,end,bottom,widths){
 sheet.showGridLines=false;
 sheet.freezePanes.unfreeze();
 sheet.tabColor='#746353';
 const all=sheet.getRange(`A1:${end}${bottom}`);
 all.format.fill='#292725';
 all.format.font={name:'Arial',size:10,color:'#F0ECE6'};
 all.format.verticalAlignment='center';
 all.format.rowHeight=24;
 sheet.getRange('A2').values=[[title]];
 sheet.getRange('A2').format.font={name:'Arial',size:16,bold:true,color:'#F0ECE6'};
 widths.forEach((width,i)=>sheet.getRangeByIndexes(0,i,bottom,1).format.columnWidth=width);
}
const first=11;
let reviewTableIndex=0;
function writeReview(main,rows,title,description){
const bottom=first+rows.length-1;
setup(main,title,'R',bottom,[9,38,16,13,25,30,17,30,17,18,35,78,36,13,17,18,22,85]);
main.getRange('A3').values=[[description]];
main.getRange('A4').values=[['Учётная дата — предложение по прямому переходу из статуса в статус. Отдельное возвращение сохраняет свою дату.']];
main.getRange('A5').values=[['Источник исходных событий: «Потенциальные смены статуса.xlsx». Строки ТВ относятся к исходной выборке РУХ_last.xlsx.']];
main.getRange('A6').values=[['Прямые исправления']];
main.getRange('D6').formulas=[[`=COUNTIF(K${first}:K${bottom},"Исправить дату")+COUNTIF(K${first}:K${bottom},"Исправить статус")+COUNTIF(K${first}:K${bottom},"Исправить статус и дату")`]];
main.getRange('F6').values=[['Предложения по учёту']];
main.getRange('I6').formulas=[[`=COUNTIF(K${first}:K${bottom},"Предлагается учётная дата")+COUNTIF(K${first}:K${bottom},"Предлагается дата приказа")`]];
main.getRange('K6').values=[['Требуют уточнения']];
main.getRange('N6').formulas=[[`=COUNTIF(K${first}:K${bottom},"Требуется проверка")+COUNTIF(K${first}:K${bottom},"Уточнить тип медицинского события")+COUNTIF(K${first}:K${bottom},"Уточнить статус по приказу")+COUNTIF(K${first}:K${bottom},"Проверить дату в приказе")+COUNTIF(K${first}:K${bottom},"Проверить противоречие")`]];
main.getRange('A7').values=[['Без изменений']];
main.getRange('D7').formulas=[[`=COUNTIF(K${first}:K${bottom},"Без изменений")`]];
main.getRange('F7').values=[['Добавить интервалы']];
main.getRange('I7').formulas=[[`=COUNTIF(K${first}:K${bottom},"Добавить промежуточное событие")`]];
main.getRange('A8').values=[['«Не определено» и пустая дата после исправления: данных недостаточно. Продовольственные поля не пересчитывались.']];
main.getRange('A9').values=[[singlePoint?'Отбор: один нумерованный пункт о человеке в приказе. Две строки событий могут относиться к одному пункту.':'Сначала прямые исправления, затем предложения по учёту, случаи для уточнения и события без изменений.']];
main.getRange(`A11:R${bottom}`).values=rows;
const headers=['Пара','ПІБ','ІПН','Строка ТВ','Операция','Подія исходная','Дата исходная','Подія после исправления','Дата после исправления','Прибуття нового интервала','Результат сверки','Обоснование решения','Куди исходное','№ приказа','Дата приказа','Пункт','Доказательства','Источник приказа'];
main.getRange('A10:R10').values=[headers];
const table=main.tables.add(`A10:R${bottom}`,true,`MovementStatusReview${++reviewTableIndex}`);
table.style='TableStyleMedium15';
main.getRange('A10:R10').format={fill:'#50463D',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:48,horizontalAlignment:'center',verticalAlignment:'center'};
const body=main.getRange(`A11:R${bottom}`);
body.format.fill='#302D29';body.format.wrapText=true;body.format.rowHeight=78;
body.format.horizontalAlignment='center';body.format.verticalAlignment='center';
body.format.borders={insideHorizontal:{style:'thin',color:'#191817'},insideVertical:{style:'thin',color:'#191817'}};
main.getRange(`B11:B${bottom}`).format.font.color='#E5AE72';
main.getRange(`C11:C${bottom}`).setNumberFormat('@');
for(const column of ['G','I','J','O'])main.getRange(`${column}11:${column}${bottom}`).setNumberFormat('dd.mm.yyyy');
body.conditionalFormats.addCustom('LEFT($K11,9)="Исправить"',{fill:'#55483A'});
body.conditionalFormats.addCustom('LEFT($K11,12)="Предлагается"',{fill:'#443C32'});
body.conditionalFormats.addCustom('OR(LEFT($K11,8)="Уточнить",LEFT($K11,9)="Проверить",$K11="Требуется проверка")',{fill:'#3D3A36',font:{color:'#F0D7AE'}});
body.conditionalFormats.addCustom('$K11="Добавить промежуточное событие"',{fill:'#504639'});
main.getRange('D6:D7').format.font.color='#E5AE72';main.getRange('I6:I7').format.font.color='#E5AE72';main.getRange('N6').format.font.color='#E5AE72';
for(let index=0;index<rows.length;index++){
 const reasonLines=Math.ceil(rows[index][11].length/75);
 const statusLines=Math.ceil(Math.max(String(rows[index][5]||'').length,String(rows[index][7]||'').length)/28);
 main.getRange(`A${index+11}:R${index+11}`).format.rowHeight=Math.max(64,Math.max(reasonLines,statusLines)*14+14);
}
}
writeReview(main,rows,singlePoint?'Смена статуса — версия 2: один пункт приказа':'Смена статуса — сверка с приказами',`В выборке ${data.decisions.length} пар событий, ${peopleCount} человек и ${orderCount} приказов А7383. ${singlePoint?`Исключено ${originalCount-data.decisions.length-unchangedPairCount} пар без подтверждения одного пункта.`:''} ${changesOnly?`Также исключено ${unchangedPairCount} пар полностью без изменений.`:''}`);
const typeSheets=[];
if(splitTypes){
 const shortNames=new Map([
  ['стаціонарне лікування','Лікування'],['відпустка для лікування','Лік.відпустка'],
  ['щорічна відпустка','Щор.відпустка'],['відрядження','Відрядження'],
  ['самовільне залишення частини','СЗЧ'],['амбулаторне влк','ВЛК'],
  ['відпустка за сімейними обставинами','Сім.відпустка'],
  ["відпустка у зв'язку з вагітністю та пологами",'Декретна відп.']
 ]);
 const groups=new Map();
 for(const decision of ordered){
  const statuses=decision.pair.slice(2,4).map(v=>v.trim().toLocaleLowerCase('uk'));
  const key=statuses.join(' → ');
  if(!groups.has(key))groups.set(key,{statuses,label:decision.pair.slice(2,4).join(' → '),ids:new Set()});
  groups.get(key).ids.add(decision.id);
 }
 for(const group of [...groups.values()].sort((a,b)=>b.ids.size-a.ids.size)){
  const name=group.statuses.map(v=>shortNames.get(v)||v).join(' → ');
  if(name.length>31)throw new Error(`Слишком длинное имя листа: ${name}`);
  const sheet=wb.worksheets.add(name);
  const selected=rows.filter(row=>group.ids.has(row[0]));
  writeReview(sheet,selected,group.label,`${group.ids.size} пар событий. Тип перехода определён по исходным статусам; исправленные события показаны рядом.`);
  typeSheets.push(name);
 }
}

const source=wb.worksheets.add('Пункты приказов');
const sourceBottom=9+sources.length;
setup(source,'Пункты приказов для проверки событий','J',sourceBottom,[12,38,13,17,18,16,115,92,85,38]);
source.getRange('A3').values=[['Текст включает применимые заголовки и продолжение персонального пункта. Даты оснований не подменяют даты событий.']];
source.getRange('A4').values=[['Номер абзаца считается среди непустых абзацев. Для поиска в Word используйте ПІБ и номер пункта.']];
source.getRange('A6').values=[[`${sources.length} относящихся к проверке фрагментов из приказов`]];
source.getRange('A6').format.font.color='#E5AE72';
source.getRange('A7').values=[['Возможное совпадение ПІБ не считается подтверждённой идентификацией человека.']];
source.getRange('A9:J9').values=[['Доказательство','ПІБ','№ приказа по ссылке','Дата в заголовке','Пункт','Абзац без пустых','Текст пункта и контекст','Підстава','Источник','Сопоставление человека']];
source.getRange(`A10:J${sourceBottom}`).values=sources;
const evidenceTable=source.tables.add(`A9:J${sourceBottom}`,true,'OrderEvidence');evidenceTable.style='TableStyleMedium15';
source.getRange('A9:J9').format={fill:'#50463D',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:48,horizontalAlignment:'center',verticalAlignment:'center'};
source.getRange(`A10:J${sourceBottom}`).format.fill='#302D29';
source.getRange(`A10:J${sourceBottom}`).format.wrapText=true;
source.getRange(`A10:J${sourceBottom}`).format.verticalAlignment='center';
source.getRange(`A10:J${sourceBottom}`).format.horizontalAlignment='center';
source.getRange(`B10:B${sourceBottom}`).format.font.color='#E5AE72';
source.getRange(`D10:D${sourceBottom}`).setNumberFormat('dd.mm.yyyy');
for(let index=0;index<sources.length;index++){
 const textLines=sources[index][6].split('\n').reduce((total,line)=>total+Math.max(1,Math.ceil(line.length/105)),0);
 const basisLines=Math.ceil(sources[index][7].length/86);
 source.getRange(`A${index+10}:J${index+10}`).format.rowHeight=Math.min(400,Math.max(80,Math.max(textLines,basisLines)*14+18));
}
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:'Сверка событий!A6:N7',include:'values,formulas',tableMaxRows:2,tableMaxCols:14,maxChars:2200})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:10},maxChars:1000})).ndjson);
for(const [sheetName,range,filename] of [['Сверка событий','A1:L15','review-preview.png'],['Пункты приказов','A1:H11','evidence-preview.png'],...(typeSheets.length?[[typeSheets[0],'A1:L15','type-preview.png']]:[])]){
 const preview=await wb.render({sheetName,range,scale:1,format:'png'});
 await fs.writeFile(`${folder}/${singlePoint?'v2-':''}${filename}`,new Uint8Array(await preview.arrayBuffer()));
}
const outputPath=nextOutputPath;
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
console.log(JSON.stringify({pairs:data.decisions.length,people:peopleCount,excluded:originalCount-data.decisions.length,rows:rows.length,evidence:sources.length,orders:orderCount,typeSheets,path:outputPath}));