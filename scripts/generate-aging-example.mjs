import { writeFile } from 'node:fs/promises';
import { createBlankWorkbook, createSheet, parseTypedCellInput, serializeWorkbook, runAgingSchedule, exportAgingReport } from '../lib/ledgerglass.js';
const workbook = createBlankWorkbook('Aging review acceptance');
workbook.id = 'aging-review-acceptance';
const sheet = workbook.sheets[0]; sheet.id = 'receivables'; sheet.name = 'Receivables';
const rows = [
 ['Amount','Currency','Due date'],
 [100,'USD','2025-04-01'],[200,'USD','2025-03-02'],[-50,'USD','2025-03-01'],
 [300,'USD','2025-01-31'],[400,'USD','2025-01-30'],[500,'USD','2025-01-01'],[600,'USD','2024-12-31'],
 [100,'USD','2025-04-02'],[50,'USD',null],[25,'USD','2025-02-30'],[1000,'EUR','2025-04-01'],
];
rows.forEach((row,r)=>row.forEach((raw,c)=>{
 const cell = typeof raw==='string' && /^\d{4}-\d{2}-\d{2}$/.test(raw) && raw!=='2025-02-30' ? parseTypedCellInput(raw) : {raw};
 sheet.cells[`${String.fromCharCode(65+c)}${r+1}`]=cell;
}));
sheet.cells.A4.note='Credit balance retained in aging totals.';
sheet.cells.E1={raw:'Type a date below to inspect ISO entry'};
sheet.cells.E2={raw:null,note:'Date entry scratch cell; source aging population is A1:C12.',tickmark:'✓',style:{bold:true}};
const hidden = createSheet('Review control'); hidden.id='review-control'; hidden.hidden=true; hidden.veryHidden=true; hidden.tabColor='#257c41'; hidden.cells={A1:{raw:'=SUM(Receivables!A2:A8)'}};workbook.sheets.push(hidden);
workbook.definedNames=[{name:'AgingPopulation',refersTo:'=Receivables!$A$1:$C$12'}];
const request={sheetId:sheet.id,range:'A1:C12',headers:true,amountColumn:0,currencyColumn:1,dateColumn:2,currency:'USD',dateBasis:'due',dateSystem:'1900',asOf:'2025-04-01',buckets:[{label:'0–30 days',from:0,to:30},{label:'31–60 days',from:31,to:60},{label:'61–90 days',from:61,to:90},{label:'91+ days',from:91,to:null}]};
const report=runAgingSchedule(workbook,request);
if(report.counts.population!==11||report.counts.aged!==7||report.totals.exactAgedAmount!=='2050'||report.totals.exactKnownSourceAmount!=='2225')throw Error('Fixture controls failed');
await writeFile(new URL('../examples/Aging-review-demo.json',import.meta.url),serializeWorkbook(workbook));
await writeFile(new URL('../examples/Aging-review-configuration.json',import.meta.url),JSON.stringify({schemaVersion:1,kind:'aging-schedule',request},null,2)+'\n');
await writeFile(new URL('../examples/Aging-review-expected-report.json',import.meta.url),exportAgingReport(workbook,report,request)+'\n');
console.log(JSON.stringify({population:report.counts,buckets:report.buckets.map(b=>({label:b.label,count:b.count,amount:b.exactAmount})),totals:report.totals}));
