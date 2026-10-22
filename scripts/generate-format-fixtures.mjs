import {build} from 'esbuild';import{writeFile,mkdir}from'node:fs/promises';import{createHash}from'node:crypto';
const root=new URL('../',import.meta.url);await mkdir(new URL('work/',root),{recursive:true});
await build({entryPoints:[new URL('tests/fixtures/alternate-formats.ts',root).pathname],outfile:new URL('work/AlternateFixtures.mjs',root).pathname,bundle:true,platform:'node',format:'esm',packages:'external'});
const {formatFixture,convertedFixture,FORMAT_ORACLE}=await import(new URL('work/AlternateFixtures.mjs',root).href),xlsx=await convertedFixture(),files={},hashes={};
for(const format of ['xls','xlsb','xlsm','xltx','xltm','ods'])files[`Format-source.${format}`]=await formatFixture(format,xlsx);
files['Format-converted-fixture.xlsx']=xlsx;
for(const[name,bytes]of Object.entries(files)){await writeFile(new URL('examples/'+name,root),bytes);hashes[name]=createHash('sha256').update(bytes).digest('hex');}
await writeFile(new URL('examples/Format-fixture-manifest.json',root),JSON.stringify({synthetic:true,oracle:FORMAT_ORACLE,externalConversionPerformed:false,nativeOfficeValidated:false,limits:'XLS/BIFF and XLSB are synthetic identity and opaque-preservation containers, not independently Office-conformant workbooks. Formula/date/style payloads are retained opaque. The paired XLSX is independently generated to test reviewed conversion workflow, not the output of a real source-application conversion.',hashes},null,2)+'\n');
await writeFile(new URL('examples/Format-fixture-data.json',root),JSON.stringify(Object.fromEntries(['xls','xlsb','xlsm','xltx','xltm','ods'].map(f=>[f,Buffer.from(files[`Format-source.${f}`]).toString('base64')])))+'\n');
console.log('Generated seven synthetic format fixtures and their feature/identity manifest.');
