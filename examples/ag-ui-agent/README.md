# AG-UI agent example

The editor's assistant panel running on an AG-UI agent. The page is `AuditWorkbook` with `createAgUiAdapter({ url: "/agent" })` as its adapter. The server answers `POST /agent` with an AG-UI 1.0 event stream. The agent is either a scripted stand-in, which needs no model and no key, or Claude through the Anthropic SDK. See [the agent guide](../../docs/AGENT-UI.md), section 6.

| File                 | What it is                                                                                                                                                   |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `ag-ui-endpoint.mjs` | Turns a responder into an AG-UI endpoint. It reads the run input, streams the reply as events, and answers 400 to a body that is not a run input.            |
| `scripted-agent.mjs` | A responder with fixed rules and no model. For a request for a total it proposes a SUM under the numbers of the selection; other requests get a text answer. |
| `claude-agent.mjs`   | A responder backed by Claude through the Anthropic SDK. It keeps each conversation in memory (see below).                                                    |
| `server.mjs`         | Serves the page and `POST /agent`. `AGENT=claude` picks Claude; the default is the scripted agent.                                                           |
| `src/main.jsx`       | The page: `AuditWorkbook` with `createAgUiAdapter({ url: "/agent" })`, and A1:B5 of a small budget.                                                          |
| `build.mjs`          | Bundles the page into `public/`.                                                                                                                             |
| `verify.mjs`         | Drives the page in a browser with the scripted agent and checks each step.                                                                                   |
| `live.mjs`           | Runs both flows against Claude from Node, without a browser. It costs money and is not part of CI.                                                           |

## Run it

From the repository root, after `npm install` and `npm run build:lib`:

```sh
node examples/ag-ui-agent/build.mjs
node examples/ag-ui-agent/server.mjs
```

Open `http://127.0.0.1:4330`, select A1:B5, open the AI assistant and ask for a total. The scripted agent answers a question in text. For "Add a total" it proposes `=SUM(B2:B4)` in B5, which you review and apply.

With Claude:

```sh
AGENT=claude node examples/ag-ui-agent/server.mjs
```

Anthropic credentials come from the environment through the Anthropic SDK: `ANTHROPIC_API_KEY`, or an `ant auth login` profile. They stay on the server and never reach the browser. The model is set by the `model` option of `createClaudeAgent`. Each request is billed.

## Check it

```sh
node examples/ag-ui-agent/verify.mjs
```

It prints four PASS lines:

```
PASS a question reaches the agent with only the approved selection, and its text answer is shown
PASS a requested change comes back as a proposal; nothing changes before review
PASS accepting the proposal puts =SUM(B2:B4) in B5 (1770)
PASS no page errors
```

The check uses Playwright's Chromium; set `CHROME_PATH` to use an installed Chrome. `npm run test:examples` runs it with the other examples.

## Live trial

```sh
node examples/ag-ui-agent/live.mjs
```

This runs the assistant flow (a question, then a change that is checked by the editor's proposal validation and applied to the workbook) and the `render_workbook` flow against Claude, from Node. Its output varies from run to run.

- It needs Anthropic credentials and costs money.
- It is not part of CI.
- The direct API path has not been run yet: no API key was available where this was built. The results recorded in [the agent guide](../../docs/AGENT-UI.md#trial-with-a-live-model) come from a different setup, a live Claude Haiku model behind the same endpoint.
- `LIVE_AGENT_MODULE` names a module whose default export is a responder to use instead of Claude.

## Writing your own agent

Any AG-UI 1.0 agent works with `createAgUiAdapter`. To serve one from Node with this example's endpoint, pass a responder to `createAgUiEndpoint(respond)`. A responder is `async (input, { signal }) => ({ text, toolCalls })`, where each tool call is `{ id, name, args }` and `args` is an object. The tools an agent may call are the frontend tools the client sent in `input.tools`: the client runs them and sends the result in its next run. `scripted-agent.mjs` is a short example of reading the selection from `input.context` and answering with a tool call.

`claude-agent.mjs` keeps each conversation in memory, per AG-UI thread, up to 200 threads, and a server restart clears them. It does this because Claude's replies must be sent back exactly as they arrived, thinking blocks included, and those do not survive a round trip through AG-UI messages.
