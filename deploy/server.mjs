/** Stateless HTTP runtime. Workbook data and service credentials never live here. */
import { createServer } from 'node:http';
import { createReadStream } from 'node:fs';
import { realpath, stat } from 'node:fs/promises';
import { resolve, sep, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { pipeline } from 'node:stream/promises';

const types = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg', '.webp': 'image/webp', '.ico': 'image/x-icon',
  '.woff': 'font/woff', '.woff2': 'font/woff2', '.ttf': 'font/ttf',
  '.wasm': 'application/wasm', '.txt': 'text/plain; charset=utf-8',
};

function integer(value, fallback, min, max, name) {
  const number = value === undefined ? fallback : Number(value);
  if (!Number.isInteger(number) || number < min || number > max || value === '')
    throw Error(`Invalid ${name}.`);
  return number;
}

function origins(value, fallback, name, allowSelf) {
  const tokens = (value ?? fallback).trim().split(/\s+/).filter(Boolean);
  if (tokens.includes("'none'")) {
    if (!allowSelf || tokens.length !== 1) throw Error(`Invalid ${name}.`);
    return tokens;
  }
  for (const token of tokens) {
    if (allowSelf && token === "'self'") continue;
    let url;
    try { url = new URL(token); } catch { throw Error(`Invalid ${name}.`); }
    if (!['https:', 'http:'].includes(url.protocol) || url.origin !== token)
      throw Error(`Use exact HTTP(S) origins in ${name}.`);
  }
  if (allowSelf && tokens.length === 0) throw Error(`Invalid ${name}.`);
  return [...new Set(tokens)];
}

export function readConfig(env = process.env) {
  return {
    root: resolve(env.STATIC_ROOT ?? fileURLToPath(new URL('../dist-container', import.meta.url))),
    host: env.HOST ?? '0.0.0.0',
    port: integer(env.PORT, 8080, 0, 65535, 'PORT'),
    shutdownMs: integer(env.SHUTDOWN_TIMEOUT_MS, 10000, 100, 60000, 'SHUTDOWN_TIMEOUT_MS'),
    frameAncestors: origins(env.FRAME_ANCESTORS, "'self'", 'FRAME_ANCESTORS', true),
    connectOrigins: origins(env.CONNECT_ORIGINS, '', 'CONNECT_ORIGINS', false),
  };
}

function acceptsGzip(header = '') {
  const encodings = header.toLowerCase().split(',').map(item => item.trim().split(';'));
  const match = encodings.find(([name]) => name === 'gzip') ?? encodings.find(([name]) => name === '*');
  if (!match) return false;
  const quality = match.slice(1).find(part => part.trim().startsWith('q='));
  const q = quality === undefined ? 1 : Number(quality.trim().slice(2));
  return Number.isFinite(q) && q > 0 && q <= 1;
}

export async function startWebServer(options = {}) {
  const config = { ...readConfig(), ...options };
  const root = await realpath(config.root);
  if (!(await stat(resolve(root, 'index.html'))).isFile()) throw Error('Missing built index.html.');
  const logger = config.logger ?? (entry => process.stdout.write(JSON.stringify(entry) + '\n'));
  const log = entry => { try { logger(entry); } catch { /* Logging must not interrupt delivery. */ } };
  let draining = false;
  const csp = [
    "default-src 'self'", "script-src 'self'", "style-src 'self' 'unsafe-inline'",
    "img-src 'self' data: blob:", "font-src 'self' data:", "worker-src 'self' blob:",
    "frame-src 'self' blob:", `connect-src 'self' ${config.connectOrigins.join(' ')}`.trim(),
    "object-src 'none'", "base-uri 'none'", "form-action 'self'",
    `frame-ancestors ${config.frameAncestors.join(' ')}`,
  ].join('; ');

  const server = createServer({ maxHeaderSize: 16 * 1024 }, async (req, res) => {
    const started = performance.now();
    // No URL, query, cookies, workbook content or authorization headers in logs.
    res.on('finish', () => {
      if (!['/healthz', '/readyz'].includes(req.url))
        log({ event: 'http', method: req.method, status: res.statusCode,
          durationMs: Math.round((performance.now() - started) * 100) / 100 });
    });
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Referrer-Policy', 'no-referrer');
    res.setHeader('Content-Security-Policy', csp);
    res.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=()');
    res.setHeader('Cache-Control', 'no-store');
    const reply = (status, text, contentType = 'text/plain; charset=utf-8') => {
      res.writeHead(status, { 'Content-Type': contentType, 'Content-Length': Buffer.byteLength(text) });
      res.end(req.method === 'HEAD' ? undefined : text);
    };
    if (!['GET', 'HEAD'].includes(req.method)) {
      res.setHeader('Allow', 'GET, HEAD');
      req.resume();
      return reply(405, 'Method not allowed.');
    }
    try {
      const rawPath = (req.url ?? '/').split('?')[0];
      if (!rawPath.startsWith('/') || rawPath.startsWith('//')) return reply(400, 'Invalid path.');
      const path = decodeURIComponent(rawPath);
      if (/[\\\x00-\x1f\x7f]/.test(path) || path.split('/').some(part => part.startsWith('.')))
        return reply(400, 'Invalid path.');
      if (path === '/healthz') return reply(200, '{"status":"alive"}', 'application/json');
      if (path === '/readyz') {
        let ready = !draining;
        try { ready &&= (await stat(resolve(root, 'index.html'))).isFile(); } catch { ready = false; }
        return reply(ready ? 200 : 503, JSON.stringify({ status: ready ? 'ready' : 'unavailable' }), 'application/json');
      }
      if (draining) { res.setHeader('Connection', 'close'); return reply(503, 'Shutting down.'); }
      const requested = resolve(root, `.${path === '/' ? '/index.html' : path}`);
      if (!requested.startsWith(root + sep)) return reply(404, 'Not found.');
      const contentType = types[extname(requested)];
      if (!contentType) return reply(404, 'Not found.');
      const file = await realpath(requested);
      if (!file.startsWith(root + sep)) return reply(404, 'Not found.');
      let selected = file;
      let info = await stat(selected);
      if (!info.isFile()) return reply(404, 'Not found.');
      res.setHeader('Vary', 'Accept-Encoding');
      if (acceptsGzip(req.headers['accept-encoding'])) {
        try {
          const zipped = await realpath(file + '.gz');
          if (zipped.startsWith(root + sep) && (await stat(zipped)).isFile()) {
            selected = zipped; info = await stat(selected); res.setHeader('Content-Encoding', 'gzip');
          }
        } catch { /* Uncompressed assets remain valid. */ }
      }
      const etag = `W/"${info.size.toString(16)}-${Math.floor(info.mtimeMs).toString(16)}-${selected === file ? 'raw' : 'gzip'}"`;
      res.setHeader('ETag', etag);
      res.setHeader('Cache-Control', path.startsWith('/assets/') && /-[\w-]{8,}\.[a-z0-9]+$/.test(path)
        ? 'public, max-age=31536000, immutable' : 'no-cache');
      if (req.headers['if-none-match'] === etag) { res.writeHead(304); res.end(); return; }
      res.writeHead(200, { 'Content-Type': contentType, 'Content-Length': info.size });
      if (req.method === 'HEAD') res.end();
      else await pipeline(createReadStream(selected), res);
    } catch (error) {
      if (res.headersSent) { res.destroy(); return; }
      if (error instanceof URIError) return reply(400, 'Invalid path.');
      if (['ENOENT', 'ENOTDIR', 'EACCES'].includes(error?.code)) return reply(404, 'Not found.');
      log({ event: 'request-error' });
      reply(500, 'Unable to serve request.');
    }
  });
  server.headersTimeout = 10000;
  server.requestTimeout = 15000;
  server.keepAliveTimeout = 5000;
  server.maxRequestsPerSocket = 1000;
  await new Promise((done, reject) => {
    server.once('error', reject);
    server.listen(config.port, config.host, () => { server.off('error', reject); done(); });
  });
  let closing;
  const close = () => closing ??= new Promise((done, reject) => {
    draining = true;
    const deadline = setTimeout(() => server.closeAllConnections(), config.shutdownMs);
    deadline.unref();
    server.close(error => { clearTimeout(deadline); error ? reject(error) : done(); });
    server.closeIdleConnections();
  });
  return { server, close, beginDrain: () => { draining = true; }, address: server.address() };
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const runtime = await startWebServer();
    process.stdout.write(JSON.stringify({ event: 'listening', port: runtime.address.port }) + '\n');
    const stop = () => { void runtime.close().catch(() => { process.exitCode = 1; }); };
    process.once('SIGTERM', stop);
    process.once('SIGINT', stop);
  } catch {
    process.stderr.write('Ledgerglass web runtime could not start. Check build output and configuration.\n');
    process.exitCode = 1;
  }
}
