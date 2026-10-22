import {readFile,writeFile,mkdir} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import assert from 'node:assert/strict';
import {build} from 'esbuild';
import {Client} from '@modelcontextprotocol/sdk/client/index.js';
import {InMemoryTransport} from '@modelcontextprotocol/sdk/inMemory.js';
import * as sdk from '../lib/ledgerglass.js';
import {createWorkbookMcpServer} from '../lib/mcp.js';
const root=new URL('../',import.meta.url);
await mkdir(new URL('work/',root),{recursive:true});
await build({entryPoints:[new URL('examples/ReportOutputFixture.ts',root).pathname],outfile:new URL('work/ReportOutputFixture.mjs',root).pathname,bundle:true,platform:'node',format:'esm',packages:'external',external:['../lib/ledgerglass.js']});
const {reportFixture,embeddedFixture}=await import(new URL('work/ReportOutputFixture.mjs',root).href);
const checks=[],check=(name,fn)=>{fn();checks.push(name);},f=reportFixture();
f.request.mode='linked';
const policy={allowedSheetIds:['source'],includeHidden:false},plan=sdk.planReportOutput(f.book,f.request,policy),published=sdk.applyReportOutputPlan(f.book,plan,policy),saved=published.reportOutputs[0],fixed=saved.snapshot;
check('captured formula values and formatting match independent arithmetic',()=>{assert.equal(fixed.rows[1].cells[1].value,1250.5);assert.equal(fixed.rows[1].cells[2].value,1250.5*.6);assert.equal(fixed.rows[1].cells[1].text,'$1,250.50');});
check('snapshot retains original workbook revision/range/time',()=>{assert.equal(fixed.workbookId,f.book.id);assert.equal(fixed.revision,0);assert.equal(fixed.request.range,'A1:C4');assert.ok(Number.isFinite(Date.parse(fixed.calculatedAt)));});
check('publication preserves source and consumes its reviewed plan',()=>{assert.deepEqual(published.sheets,f.book.sheets);assert.throws(()=>sdk.applyReportOutputPlan(f.book,plan,policy));});
check('current linked read recomputes authorized values',()=>assert.equal(sdk.readTrustedReportOutput(published,saved.id,policy).rows[1].cells[2].value,750.3));
const stale=structuredClone(published);stale.sheets[0].cells.B2.raw=7;
check('source edits stale the link but preserve historical snapshot bytes',()=>{assert.equal(sdk.inspectReportOutput(stale,saved.id,policy).status,'stale');assert.deepEqual(stale.reportOutputs[0].snapshot,fixed);assert.throws(()=>sdk.readTrustedReportOutput(stale,saved.id,policy));});
const refreshed=sdk.applyReportOutputPlan(stale,sdk.planReportOutput(stale,saved.snapshot.request,policy,saved.id),policy);
check('explicit reviewed refresh increments version and changes only saved snapshot',()=>{assert.equal(refreshed.reportOutputs[0].version,2);assert.equal(refreshed.reportOutputs[0].snapshot.rows[1].cells[1].value,7);assert.equal(fixed.rows[1].cells[1].value,1250.5);});
check('unauthorized, forged and read-only plans cannot publish',()=>{assert.throws(()=>sdk.captureReportSnapshot(f.book,f.request,{allowedSheetIds:[]}));assert.throws(()=>sdk.applyReportOutputPlan(f.book,structuredClone(sdk.planReportOutput(f.book,f.request,policy)),policy));assert.throws(()=>sdk.applyReportOutputPlan(f.book,sdk.planReportOutput(f.book,f.request,policy),{...policy,readOnly:true}));});
const filter={range:'A1:C4',columns:[{column:0,values:['East']}]},visible=sdk.captureReportSnapshot(f.book,{...f.request,name:'East visible schedule',population:'visible',filter},policy);
check('visible filter retains expected rows without excluded values',()=>{assert.deepEqual(visible.rows.map(r=>r.row),[1,2,4]);assert.deepEqual(visible.excludedRows,[3]);assert.ok(!JSON.stringify(visible).includes('West — withheld'));assert.deepEqual(visible.request.filter,filter);});
check('full population retains values and explicit filter context',()=>{const s=sdk.captureReportSnapshot(f.book,{...f.request,filter},policy);assert.equal(s.rows.length,4);assert.equal(s.rows[2].cells[1].value,2000);assert.ok(s.request.filter);});
check('resource and unsupported-calculation refusal',()=>{assert.throws(()=>sdk.captureReportSnapshot(f.book,{...f.request,range:'A1:XFD1048576'},policy));const b=structuredClone(f.book);b.sheets[0].cells.C2.raw='=RAND()';assert.throws(()=>sdk.captureReportSnapshot(b,f.request,policy));});
check('hidden dependency permission is enforced',()=>{const b=structuredClone(f.book);b.sheets[0].hiddenRows=[1];assert.throws(()=>sdk.captureReportSnapshot(b,f.request,policy));});
const files={};for(const [suffix,exporter]of [['svg',sdk.reportSnapshotSvg],['html',sdk.reportSnapshotHtml],['json',sdk.reportSnapshotJson]]){
  const result=exporter(visible);files[`examples/Report-output-visible.${suffix}`]=result.content;
  check(`${suffix} export omits excluded values and records source context`,()=>{assert.ok(result.content.includes('A1:C4'));assert.ok(!result.content.includes('West — withheld'));});
}
check('HTML is self-contained with a restrictive content policy and accessible table',()=>{const h=files['examples/Report-output-visible.html'];assert.match(h,/Content-Security-Policy/);assert.match(h,/<table/);assert.ok(!/<script/i.test(h));assert.match(h,/max-width:100%/);});
for(const [id,range]of [['chart','A1:B4'],['picture','E22'],['shape','E16']]){
  const s=sdk.captureReportSnapshot(f.book,{...f.request,name:`Report ${id}`,kind:'object',objectId:id,range},policy),v=sdk.reportSnapshotSvg(s);files[`examples/Report-output-${id}.svg`]=v.content;
  check(`${id} snapshot renders without remote image or executable content`,()=>{assert.ok(v.content.startsWith('<svg'));assert.ok(!/<script|<foreignObject|href="https?:/i.test(v.content));assert.equal(s.object.id,id);});
}
const native=await sdk.importFile(new File([sdk.serializeWorkbook(published)],'report.json'));
check('actual native file-worker reopen preserves report definitions and values',()=>{assert.deepEqual(native.workbook.reportOutputs,published.reportOutputs);assert.equal(sdk.readTrustedReportOutput(native.workbook,saved.id,policy).rows[1].cells[2].value,750.3);});
const objectBytes=embeddedFixture(),inspection=await sdk.inspectAutomationFile(new File([objectBytes],'objects.xlsx')),inventoried=sdk.retainEmbeddedObjectInventory(published,inspection);
check('isolated inspection retains five inert object references',()=>{const i=inventoried.embeddedObjectInventories[0];assert.equal(i.items.length,5);assert.equal(i.execution,'never');assert.ok(i.items.every(v=>v.execution==='never'));assert.ok(i.items.some(v=>/missing/.test(v.path)));});
await assert.rejects(sdk.importFile(new File([objectBytes],'objects.xlsx')));checks.push('normal file processing still rejects the executable container');
const exchange=await sdk.exportXlsxWithWarnings(inventoried);
check('actual XLSX worker emits per-item dropped losses and native report warning',()=>{assert.equal(exchange.report.embeddedObjectLosses.length,5);assert.ok(exchange.report.embeddedObjectLosses.every(i=>i.disposition==='dropped'&&i.execution==='never'));assert.match(exchange.warnings.join(' '),/report outputs.*omitted from XLSX/);});
const derived=await sdk.importFile(new File([exchange.blob],'derived.xlsx'));
check('derived XLSX reopens with supported cells and no executable report metadata',()=>{assert.equal(derived.workbook.sheets[0].cells.B2.raw,1250.5);assert.equal(derived.workbook.reportOutputs,undefined);assert.equal(derived.workbook.embeddedObjectInventories,undefined);});
const nativeInventory=await sdk.importFile(new File([sdk.serializeWorkbook(inventoried)],'inventory.json'));
check('native file-worker reopen preserves exact inert inventory',()=>assert.deepEqual(nativeInventory.workbook.embeddedObjectInventories,inventoried.embeddedObjectInventories));
const server=createWorkbookMcpServer({getWorkbook:()=>inventoried,getSelection:()=>({sheetId:'source',range:'A1:A1'}),queueProposal:()=>{throw Error('Unexpected mutation');}}),client=new Client({name:'report-local-acceptance',version:'1'}),[a,b]=InMemoryTransport.createLinkedPair();await server.connect(b);await client.connect(a);
try{const tools=(await client.listTools()).tools,caps=(await client.callTool({name:'workbook_get_capabilities',arguments:{}})).structuredContent,context=JSON.stringify(await client.callTool({name:'workbook_get_context',arguments:{}}));check('in-memory MCP advertises reuse with unchanged nine-tool authority',()=>{assert.equal(tools.length,9);assert.equal(caps.reusableReports.engine,sdk.REPORT_OUTPUT_ENGINE);});check('MCP approved context excludes report cache and inventory',()=>{for(const token of ['1250.5','sourceHash','Regional schedule','opaque.dat'])assert.ok(!context.includes(token));});}finally{await client.close();await server.close();}
for(const[name,book]of [['source',f.book],['result',published],['refreshed',refreshed],['inventory',inventoried]])files[`examples/Report-output-${name}.json`]=sdk.serializeWorkbook(book);
files['examples/Report-output-losses.json']=JSON.stringify(exchange.report,null,2)+'\n';
for(const[path,content]of Object.entries(files))await writeFile(new URL(path,root),content);
const hashes={};for(const path of ['lib/ledgerglass.js','lib/ledgerglass.css','lib/mcp.js','lib/file-worker.js','lib/source-gateway.js','lib/source-connectors.js','lib/sqlite-source-worker.cjs','Ledgerglass.html','dist/report-outputs.html',...Object.keys(files)])hashes[path]=createHash('sha256').update(await readFile(new URL(path,root))).digest('hex');
const proof={requirement:'AW-063',engine:sdk.REPORT_OUTPUT_ENGINE,checksPassed:checks.length,checks,hashes,limits:sdk.REPORT_OUTPUT_LIMITS,environment:{node:process.version,platform:process.platform,architecture:process.arch},actualIO:{nativeJsonFileWorker:true,objectInspectionFileWorker:true,xlsxExportFileWorker:true,xlsxReopenFileWorker:true,inMemoryMcpClient:true},external:{nativeOfficeExecuted:false,externalLLMTested:false,independentHostAcceptance:false},expected:{sourceValue:1250.5,calculatedMargin:750.3,visibleRows:[1,2,4],excludedRows:[3],refreshedValue:7,refreshedVersion:2,objectLosses:5},boundary:'Reviewed local links and inert snapshots/inventory; no native Office OLE execution, automatic remote refresh, real external LLM or independent production-host acceptance.'};
await writeFile(new URL('examples/Report-output-verification.json',root),JSON.stringify(proof,null,2)+'\n');console.log(JSON.stringify({checksPassed:checks.length,expected:proof.expected}));
