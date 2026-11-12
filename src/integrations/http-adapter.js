/**
 * Assistant adapter that asks your own server for a proposal.
 *
 * The editor calls `propose(prompt, context, signal, request)` and expects a proposal object
 * (`{ id, workbookId, revision, explanation, citations, patches }`); it validates the reply against
 * the approved selection before showing it. This adapter POSTs those arguments as JSON to one
 * endpoint. Keep model provider keys on that server, never in the browser.
 *
 * Written for the rebuilt library; the original adapter's wire format was not recovered.
 */
function createHttpAdapter({
  endpoint,
  name = "Workbook assistant",
  chat = false,
  headers = {},
  fetch: fetchImplementation = globalThis.fetch,
} = {}) {
  if (typeof endpoint !== "string" || !endpoint)
    throw new TypeError("createHttpAdapter needs an endpoint.");
  return {
    name,
    ...(chat ? { chat: { history: true, attachments: false } } : {}),
    async propose(prompt, context, signal, request) {
      const response = await fetchImplementation(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json", ...headers },
        body: JSON.stringify({ prompt, context, request }),
        signal,
        credentials: "same-origin",
      });
      if (!response.ok)
        throw new Error(
          `The assistant service answered ${response.status}. No change was proposed.`,
        );
      return response.json();
    },
  };
}

export { createHttpAdapter };
