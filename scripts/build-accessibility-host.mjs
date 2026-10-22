import {build} from 'esbuild';
import {writeFile} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
const root=new URL('../',import.meta.url);
const result=await build({absWorkingDir:fileURLToPath(root),entryPoints:['examples/AccessibilityHost.tsx'],outfile:'accessibility.js',bundle:true,write:false,minify:true,platform:'browser',format:'esm',target:'es2022',jsx:'automatic',define:{'process.env.NODE_ENV':'"production"'},logLevel:'warning'});
const script=result.outputFiles.find(f=>f.path.endsWith('.js')).text,css=result.outputFiles.find(f=>f.path.endsWith('.css')).text;
const host='html,body,#root{height:100%;margin:0}body{font:13px Arial,sans-serif;color:#14234b}.a11y-host{height:100%;display:flex;flex-direction:column}.a11y-host>header{padding:10px 16px;background:#f5f7fd;border-bottom:1px solid #8794b0}.a11y-host h1{font-size:18px;margin:0 0 6px}.a11y-host p{margin:5px 0}.a11y-actions{display:flex;gap:8px;flex-wrap:wrap;align-items:center}.a11y-actions button,.a11y-actions select{font:inherit;background:white;border:1px solid #8794b0;border-radius:4px;padding:6px;color:#14234b}.a11y-host output{display:block;margin-top:6px;font:12px monospace}.a11y-host pre{max-height:180px;overflow:auto;white-space:pre-wrap}.a11y-workbook{flex:1;min-height:0}.a11y-workbook.enlarged{zoom:2;min-height:0}.a11y-workbook .lg-app{min-height:0;height:100%}';
// CSS magnification does not resize viewport units as native browser zoom does.
// Constrain fixed dialogs to the equivalent viewport inside this test harness.
const enlarged='.a11y-workbook.enlarged [role="dialog"][aria-modal="true"]{max-height:calc(50dvh - 24px);max-width:calc(50dvw - 24px)}';
await writeFile(new URL('dist/accessibility.html',root),`<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Workbook accessibility and localization</title><style>${(css+host+enlarged).replaceAll('</style','<\\/style')}</style></head><body><div id="root"></div><script type="module">${script.replaceAll('</script','<\\/script')}</script></body></html>`);
console.log('Created dist/accessibility.html from the compiled SDK with axe-core 4.11.1.');
