import { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { z } from "zod/v3";
import { A2uiSurface, createComponentImplementation } from "@a2ui/react";
import { Catalog, CommonSchemas, MessageProcessor } from "@a2ui/web_core/v1_0";
import { basicCatalog } from "@a2ui/web_core/catalogs/basic/v1";
import { HttpAgent } from "@ag-ui/client";
import {
  AgentWorkbook,
  createWorkbookCatalog,
  createWorkbookToolSubscriber,
  renderWorkbookTool,
  WORKBOOK_CATALOG_ID,
} from "../../lib/ledgerglass.js";
import "../../lib/ledgerglass.css";

const mode = new URLSearchParams(location.search).get("mode");
window.log = [];

function A2uiHost() {
  const [processor] = useState(() => {
    const workbookCatalog = createWorkbookCatalog({
      Catalog,
      CommonSchemas,
      createComponentImplementation,
      z,
    });
    window.catalogJson = workbookCatalog.toJson?.();
    const p = new MessageProcessor([basicCatalog, workbookCatalog], (action) =>
      window.log.push({ action }),
    );
    window.processor = p;
    window.capabilities = p.getRendererCapabilities({ versions: ["1.0"] });
    window.workbookCatalogId = WORKBOOK_CATALOG_ID;
    return p;
  });
  const [surfaces, setSurfaces] = useState([]);
  useEffect(() => {
    const sync = () =>
      setSurfaces(Array.from(processor.model.surfacesMap.values()));
    const a = processor.onSurfaceCreated(sync),
      b = processor.onSurfaceDeleted(sync);
    window.send = (messages) => {
      try {
        processor.processMessages(messages);
        return "ok";
      } catch (e) {
        return String((e && e.message) || e);
      }
    };
    window.dataModel = (id) =>
      processor.model.surfacesMap.get(id)?.dataModel.get("/");
    window.ready = true;
    return () => {
      a.unsubscribe();
      b.unsubscribe();
    };
  }, [processor]);
  return (
    <div>
      {surfaces.map((s) => (
        <A2uiSurface key={s.id} surface={s} />
      ))}
    </div>
  );
}

function AgUiHost() {
  const [calls, setCalls] = useState([]);
  useEffect(() => {
    const agent = new HttpAgent({ url: "/agent" });
    const pending = [];
    agent.subscribe(
      createWorkbookToolSubscriber({
        onWorkbook: (call) => {
          pending.push(call.result);
          window.log.push({ shown: call.toolCallId, result: call.result });
          setCalls((c) => [...c, call]);
        },
        onInvalid: (call) => {
          pending.push(call.result);
          window.log.push({ invalid: call.problems, result: call.result });
        },
      }),
    );
    window.run = async () => {
      for (const m of pending.splice(0)) agent.addMessage(m);
      try {
        const r = await agent.runAgent({ tools: [renderWorkbookTool] });
        return { ok: true, newMessages: r.newMessages.length };
      } catch (e) {
        return { ok: false, error: String((e && e.message) || e) };
      }
    };
    window.ready = true;
  }, []);
  return (
    <div>
      {calls.map((c) => (
        <AgentWorkbook
          key={c.toolCallId}
          spec={c.spec}
          workbookId={c.toolCallId}
          height={420}
        />
      ))}
    </div>
  );
}
createRoot(document.getElementById("host")).render(
  mode === "agui" ? <AgUiHost /> : <A2uiHost />,
);
