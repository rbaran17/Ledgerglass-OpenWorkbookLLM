/** Runs only against a disposable, uniquely named Compose project. Requires Docker. */
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { writeFile } from 'node:fs/promises';
import assert from 'node:assert/strict';
const exec = promisify(execFile);
const project = `ledgerglass-check-${process.pid}-${Date.now()}`;
const env = { ...process.env, LEDGERGLASS_PORT: '0' };
const run = async args => (await exec('docker', args, {
  env, maxBuffer: 10 * 1024 * 1024, timeout: 20 * 60 * 1000,
})).stdout.trim();
const compose = (...args) => run(['compose', '-p', project, '-f', 'compose.yaml', ...args]);
const checks = [];
const record = name => { checks.push(name); console.log(`PASS ${name}`); };
await run(['info', '--format', '{{.ServerVersion}}']);
try {
  const config = JSON.parse(await compose('config', '--format', 'json'));
  assert.equal(config.services.web.build.target, 'web');
  assert.equal(config.services.web.read_only, true);
  assert.equal(config.services.web.ports, undefined);
  record('Compose configuration');
  await compose('up', '--build', '-d', '--scale', 'web=3', '--wait', '--wait-timeout', '180');
  const ids = (await compose('ps', '-q', 'web')).split(/\s+/);
  assert.equal(ids.length, 3);
  const inspections = JSON.parse(await run(['inspect', ...ids]));
  for (const container of inspections) {
    assert.ok(['node', '1000', '1000:1000'].includes(container.Config.User));
    assert.equal(container.HostConfig.ReadonlyRootfs, true);
    assert.equal(container.State.Health.Status, 'healthy');
  }
  record('Three healthy non-root read-only replicas');
  const address = await compose('port', 'gateway', '8080');
  const url = `http://${address}`;
  const response = await fetch(`${url}/?embedded`);
  assert.equal(response.status, 200);
  assert.match(response.headers.get('content-security-policy'), /frame-ancestors/);
  const html = await response.text();
  const asset = html.match(/src="([^"]+\.js)"/)?.[1];
  assert.ok(asset, 'HTML entry module');
  const assetResponse = await fetch(new URL(asset, url + '/'));
  assert.equal(assetResponse.status, 200);
  assert.equal(assetResponse.headers.get('content-encoding'), 'gzip');
  record('Load balancer serves application, CSP and compressed assets');
  const responses = await Promise.all(Array.from({ length: 60 }, () => fetch(`${url}/readyz`)));
  assert.ok(responses.every(item => item.status === 200));
  await Promise.all(responses.map(item => item.arrayBuffer()));
  record('Concurrent requests');
  await compose('up', '-d', '--scale', 'web=2', '--wait', '--wait-timeout', '90');
  assert.equal((await compose('ps', '-q', 'web')).split(/\s+/).length, 2);
  // The resolver refreshes every 5 seconds; retries permit only a bounded recovery interval.
  const deadline = Date.now() + 20000;
  let recovered = false;
  while (Date.now() < deadline) {
    try { const result = await fetch(`${url}/readyz`); await result.arrayBuffer(); if (result.ok) { recovered = true; break; } } catch {}
    await new Promise(done => setTimeout(done, 500));
  }
  assert.ok(recovered);
  record('Scale-down remains reachable');
  const mcpImage = `${project}-mcp`;
  try {
    await run(['build', '--target', 'mcp', '-t', mcpImage, '.']);
    const help = await exec('docker', ['run', '--rm', '--read-only', '--network', 'none',
      '--cap-drop', 'ALL', '--security-opt', 'no-new-privileges:true', mcpImage, '--help'], { maxBuffer: 1024 * 1024 });
    assert.match(help.stderr, /queues proposals only/);
    record('Isolated stdio MCP image starts');
  } finally { await run(['image', 'rm', mcpImage]).catch(() => {}); }
  await writeFile('deploy/container-verification.json', JSON.stringify({
    executedAt: new Date().toISOString(), result: 'passed', checks,
    productionDeployment: 'not run', externalLlm: 'not run', clusterAutoscaling: 'not run',
  }, null, 2) + '\n');
} finally {
  await compose('down', '--volumes', '--remove-orphans');
}
