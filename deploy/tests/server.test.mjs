import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, rm, symlink, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { request } from 'node:http';
import { brotliCompressSync, brotliDecompressSync, gzipSync, gunzipSync } from 'node:zlib';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { readConfig, startWebServer } from '../server.mjs';

async function setup(t, extra = {}) {
  const dir = await mkdtemp(join(tmpdir(), 'ledgerglass-web-'));
  const root = join(dir, 'public');
  await mkdir(join(root, 'assets'), { recursive: true });
  await writeFile(join(root, 'index.html'), '<!doctype html><title>Workbook</title>');
  await writeFile(join(root, 'assets/app-abcdef12.js'), 'console.log("Workbook");');
  await writeFile(join(root, 'assets/app-abcdef12.js.gz'), gzipSync('console.log("Workbook");'));
  await writeFile(join(dir, 'secret.txt'), 'private');
  await symlink(join(dir, 'secret.txt'), join(root, 'outside.txt'));
  const logs = [];
  const runtime = await startWebServer({ root, host: '127.0.0.1', port: 0, logger: item => logs.push(item), ...extra });
  t.after(async () => { await runtime.close(); await rm(dir, { recursive: true, force: true }); });
  const get = (path, options = {}) => new Promise((done, reject) => {
    const req = request({ host: '127.0.0.1', port: runtime.address.port, path, ...options }, res => {
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => done({ status: res.statusCode, headers: res.headers, body: Buffer.concat(chunks) }));
    });
    req.on('error', reject); req.end(options.body);
  });
  return { dir, root, runtime, logs, get };
}

test('serves the workbook and embedded query without application state or cookies', async t => {
  const { get } = await setup(t);
  const result = await get('/?embedded');
  assert.equal(result.status, 200);
  assert.match(result.body.toString(), /Workbook/);
  assert.equal(result.headers['set-cookie'], undefined);
  assert.equal(result.headers['cache-control'], 'no-cache');
});
test('health and readiness have separate drain behavior', async t => {
  const { get, runtime } = await setup(t);
  assert.equal((await get('/readyz')).status, 200);
  runtime.beginDrain();
  assert.equal((await get('/healthz')).status, 200);
  assert.equal((await get('/readyz')).status, 503);
  assert.equal((await get('/')).status, 503);
});
test('readiness fails when the built application disappears', async t => {
  const { get, root } = await setup(t);
  await rm(join(root, 'index.html'));
  assert.equal((await get('/readyz')).status, 503);
});
test('HEAD returns headers without a body', async t => {
  const { get } = await setup(t);
  const response = await get('/', { method: 'HEAD' });
  assert.equal(response.status, 200);
  assert.ok(Number(response.headers['content-length']) > 0);
  assert.equal(response.body.length, 0);
});
test('hashed assets are cacheable and matching conditional requests are 304', async t => {
  const { get } = await setup(t);
  const response = await get('/assets/app-abcdef12.js');
  assert.match(response.headers['cache-control'], /immutable/);
  const cached = await get('/assets/app-abcdef12.js', { headers: { 'If-None-Match': response.headers.etag } });
  assert.equal(cached.status, 304);
  assert.equal(cached.body.length, 0);
});
test('gzip is byte-correct and q=0 explicitly opts out', async t => {
  const { get } = await setup(t);
  const response = await get('/assets/app-abcdef12.js', { headers: { 'Accept-Encoding': 'gzip' } });
  assert.equal(response.headers['content-encoding'], 'gzip');
  assert.equal(Number(response.headers['content-length']), response.body.length);
  assert.equal(gunzipSync(response.body).toString(), 'console.log("Workbook");');
  const plain = await get('/assets/app-abcdef12.js', { headers: { 'Accept-Encoding': '*;q=1, gzip;q=0' } });
  assert.equal(plain.headers['content-encoding'], undefined);
  assert.notEqual(plain.headers.etag, response.headers.etag);
});
test('brotli is preferred when offered and falls back to gzip', async t => {
  const { get, root } = await setup(t);
  const source = 'console.log("Workbook");';
  const path = '/assets/app-abcdef12.js';
  const withoutSibling = await get(path, { headers: { 'Accept-Encoding': 'gzip, br' } });
  assert.equal(withoutSibling.headers['content-encoding'], 'gzip');
  await writeFile(join(root, 'assets/app-abcdef12.js.br'), brotliCompressSync(source));
  const response = await get(path, { headers: { 'Accept-Encoding': 'gzip, deflate, br' } });
  assert.equal(response.headers['content-encoding'], 'br');
  assert.equal(Number(response.headers['content-length']), response.body.length);
  assert.equal(brotliDecompressSync(response.body).toString(), source);
  assert.notEqual(response.headers.etag, withoutSibling.headers.etag);
  const refused = await get(path, { headers: { 'Accept-Encoding': 'br;q=0, gzip' } });
  assert.equal(refused.headers['content-encoding'], 'gzip');
  assert.equal((await get(path + '.br')).status, 404);
});
test('responses isolate the workbook from other origins', async t => {
  const { get } = await setup(t);
  for (const path of ['/', '/assets/app-abcdef12.js', '/missing.js']) {
    const { headers } = await get(path);
    assert.equal(headers['cross-origin-opener-policy'], 'same-origin');
    assert.equal(headers['cross-origin-resource-policy'], 'same-origin');
  }
});
test('shutdown does not wait for the keep-alive timeout', async t => {
  const { get, runtime } = await setup(t);
  await get('/');
  const started = performance.now();
  await runtime.close();
  assert.ok(performance.now() - started < 1000);
});
test('CSP permits browser workers and explicitly approved embedding origins', async t => {
  const { get } = await setup(t, { frameAncestors: ["'self'", 'https://host.example'], connectOrigins: ['https://api.example'] });
  const { headers } = await get('/');
  assert.match(headers['content-security-policy'], /worker-src 'self' blob:/);
  assert.match(headers['content-security-policy'], /frame-ancestors 'self' https:\/\/host.example/);
  assert.match(headers['content-security-policy'], /connect-src 'self' https:\/\/api.example/);
  assert.equal(headers['x-frame-options'], undefined);
  assert.equal(headers['cross-origin-resource-policy'], 'cross-origin');
  assert.equal(headers['x-content-type-options'], 'nosniff');
});
test('rejects unsupported methods and never exposes API or source files', async t => {
  const { get } = await setup(t);
  assert.equal((await get('/', { method: 'POST', body: '{"workbook":"sensitive"}' })).status, 405);
  for (const path of ['/api/workbooks', '/package.json', '/src/main.tsx', '/missing.js'])
    assert.equal((await get(path)).status, 404);
});
test('rejects traversal, malformed escapes, dotfiles and symlink escapes', async t => {
  const { get } = await setup(t);
  for (const path of ['/../secret.txt', '/%2e%2e/secret.txt', '/.env', '/%ZZ', '/assets%5csecret.txt', '/%00', '//outside'])
    assert.equal((await get(path)).status, 400, path);
  assert.equal((await get('/outside.txt')).status, 404);
});
test('logs contain no URLs, authorization, cookies or queries', async t => {
  const { get, logs } = await setup(t);
  await get('/?workbook=private-id', { headers: { authorization: 'Bearer secret', cookie: 'session=secret' } });
  // The entry is written when the response finishes, which can follow the client's last byte.
  for (let attempt = 0; attempt < 100 && logs.length === 0; attempt++) await new Promise(done => setTimeout(done, 10));
  const serialized = JSON.stringify(logs);
  assert.equal(logs.length, 1);
  assert.doesNotMatch(serialized, /private-id|secret|workbook=/);
  assert.equal(logs[0].status, 200);
});
test('invalid port, shutdown deadline and CSP origin injection fail configuration', () => {
  for (const config of [{ PORT: 'bad' }, { PORT: '' }, { SHUTDOWN_TIMEOUT_MS: '0' },
    { FRAME_ANCESTORS: '*' }, { FRAME_ANCESTORS: "'none' https://host.example" },
    { CONNECT_ORIGINS: 'https://host.example; script-src *' }, { CONNECT_ORIGINS: 'https://host.example/path' }])
    assert.throws(() => readConfig(config));
  assert.deepEqual(readConfig({ FRAME_ANCESTORS: "'none'" }).frameAncestors, ["'none'"]);
});
test('two replicas return identical content without sticky sessions', async t => {
  const a = await setup(t), b = await setup(t);
  const responses = await Promise.all(Array.from({ length: 40 }, (_, i) => (i % 2 ? a : b).get('/')));
  assert.ok(responses.every(response => response.status === 200 && response.body.equals(responses[0].body)));
});
test('asset preparation creates a verifiable manifest and gzip siblings', async t => {
  const { root } = await setup(t);
  await rm(join(root, 'outside.txt'));
  const child = spawn(process.execPath, ['deploy/prepare-assets.mjs', root], { stdio: 'pipe' });
  let error = ''; child.stderr.on('data', chunk => { error += chunk; });
  const [code] = await once(child, 'exit');
  assert.equal(code, 0, error);
  const manifest = JSON.parse(await readFile(join(root, 'asset-manifest.json'), 'utf8'));
  assert.equal(manifest.files.length, 2);
  const html = await readFile(join(root, 'index.html'));
  assert.deepEqual(gunzipSync(await readFile(join(root, 'index.html.gz'))), html);
  assert.deepEqual(brotliDecompressSync(await readFile(join(root, 'index.html.br'))), html);
});
test('SIGTERM exits cleanly from the actual executable runtime', async t => {
  const { root } = await setup(t);
  const child = spawn(process.execPath, ['deploy/server.mjs'], {
    env: { ...process.env, PORT: '0', HOST: '127.0.0.1', STATIC_ROOT: root }, stdio: 'pipe',
  });
  t.after(() => { if (child.exitCode === null) child.kill('SIGKILL'); });
  await once(child.stdout, 'data');
  const exited = once(child, 'exit'); child.kill('SIGTERM');
  assert.deepEqual(await exited, [0, null]);
});
