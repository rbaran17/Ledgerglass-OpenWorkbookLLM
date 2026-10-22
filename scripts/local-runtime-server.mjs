/** Reproducible local-runtime harness. Build SDK first. Ctrl+C disconnects all hosting. */
import { createServer } from 'node:http';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
const root=new URL('../',import.meta.url),dir=new URL('work/local-runtime/',root);
await mkdir(dir,{recursive:true});
await build({entryPoints:[fileURLToPath(new URL('examples/LocalRuntimeHost.tsx',root))],outfile:fileURLToPath(new URL('host.js',dir)),bundle:true,format:'iife',jsx:'automatic',minify:true,define:{'process.env.NODE_ENV':'"production"'}});
const js=await readFile(new URL('host.js',dir),'utf8'),css=await readFile(new URL('host.css',dir),'utf8');
const html='<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Local runtime acceptance</title><style>'+css.replaceAll('</style','<\\/style')+'</style><body style="margin:0"><div id="root"></div><script>'+js.replaceAll('</script','<\\/script')+'</script></body></html>';
await writeFile(new URL('host.html',dir),html);
const port=Number(process.argv.find(a=>a.startsWith('--port='))?.split('=')[1]??0);
const server=createServer(async(req,res)=>{
 if(req.url==='/assistant'&&req.method==='POST'){
  try{let body='';for await(const chunk of req){body+=chunk;if(body.length>1000000)throw Error('size');}const {context}=JSON.parse(body);res.writeHead(200,{'Content-Type':'application/json'});res.end(JSON.stringify({id:'host-'+crypto.randomUUID(),workbookId:context.workbookId,revision:context.revision,explanation:'Connected test service responded. No model was invoked.',citations:context.cells.slice(0,1).map(c=>c.address),patches:[]}));}catch{res.writeHead(400).end();}return;
 }
 if(req.url==='/favicon.ico'){res.writeHead(204).end();return;}
 res.writeHead(200,{'Content-Type':'text/html; charset=utf-8','Cache-Control':'no-store'});res.end(html);
});
server.listen(port,'127.0.0.1',()=>console.log(JSON.stringify({url:'http://127.0.0.1:'+server.address().port,procedure:'Load the host, then stop this server. Existing local controls remain available; Host test service reports unavailable. Restart with the same --port to verify retry.'})));
