/**
 * External source adapter for the editor's `externalSourceAdapter` prop: it sends each
 * operation to your server's source gateway and returns the reply. The editor validates every
 * reply itself, so this adapter only carries requests and maps failures to protocol errors.
 *
 * The wire format is the request shape found in the recovered editor:
 * `{ operation: "catalog" }` or `{ operation: "authorize" | "refresh", sourceId }`, answered with
 * JSON, or with `{ error: { code } }` and a non-2xx status.
 */
import { ExternalSourceError } from "../app/03-workbook-model.js";

const knownCodes = new Set([
  "unavailable",
  "refused",
  "expired",
  "schema",
  "limit",
  "timeout",
  "cancelled",
  "invalid",
  "stale",
]);

function createHttpSourceAdapter({
  endpoint,
  headers = {},
  fetch: fetchImplementation = globalThis.fetch,
  credentials = "same-origin",
} = {}) {
  if (typeof endpoint !== "string" || !endpoint)
    throw new TypeError("createHttpSourceAdapter needs an endpoint.");

  async function send(body, signal) {
    let response;
    try {
      response = await fetchImplementation(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json", ...headers },
        body: JSON.stringify(body),
        signal,
        credentials,
      });
    } catch (error) {
      throw new ExternalSourceError(
        error?.name === "AbortError" ? "cancelled" : "unavailable",
      );
    }
    let reply;
    try {
      reply = await response.json();
    } catch {
      throw new ExternalSourceError(response.ok ? "invalid" : "unavailable");
    }
    if (response.ok) return reply;
    const code = reply?.error?.code;
    throw new ExternalSourceError(
      knownCodes.has(code)
        ? code
        : response.status === 401 || response.status === 403
          ? "refused"
          : "unavailable",
    );
  }

  return {
    catalog: ({ signal } = {}) => send({ operation: "catalog" }, signal),
    authorize: (sourceId, { signal } = {}) =>
      send({ operation: "authorize", sourceId }, signal),
    refresh: (sourceId, { signal } = {}) =>
      send({ operation: "refresh", sourceId }, signal),
  };
}

export { createHttpSourceAdapter };
