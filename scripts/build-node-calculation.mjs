import {build} from 'esbuild';
const root=new URL('../',import.meta.url);
await build({entryPoints:[new URL('src/workers/calculation-node-entry.ts',root).pathname],outfile:new URL('lib/calculation-worker.cjs',root).pathname,bundle:true,platform:'node',target:'node22',format:'cjs',minify:true,logLevel:'warning'});
await build({entryPoints:[new URL('src/integrations/node-calculation.ts',root).pathname],outfile:new URL('lib/node-calculation.js',root).pathname,bundle:true,platform:'node',target:'node22',format:'esm',logLevel:'warning'});
