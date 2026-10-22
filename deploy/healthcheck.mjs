try {
  const response = await fetch(`http://127.0.0.1:${process.env.PORT ?? 8080}/readyz`, { signal: AbortSignal.timeout(2500) });
  if (!response.ok || (await response.json()).status !== 'ready') process.exitCode = 1;
} catch { process.exitCode = 1; }
