import { useMemo } from 'react';
import {
  AuditWorkbook, createHttpAdapter, createHttpExternalSourceAdapter,
  type Workbook, type WorkbookStoreLike, type PersistenceState,
} from '@ledgerglass/workbook';
import '@ledgerglass/workbook/style.css';

/** The host loads and authorizes the workbook before mounting this component. */
export function WorkbookWindow(props: {
  initialWorkbook: Workbook;
  store: WorkbookStoreLike;
  actor: string;
  onSaveState: (state: PersistenceState) => void;
}) {
  const assistant = useMemo(() => createHttpAdapter({
    endpoint: '/api/workbook-assistant', chat: true,
  }), []);
  const sources = useMemo(() => createHttpExternalSourceAdapter({
    endpoint: '/api/workbook-sources',
  }), []);
  return <div style={{ height: 'calc(100dvh - 160px)', minHeight: 420 }}>
    <AuditWorkbook
      key={props.initialWorkbook.id}
      initialWorkbook={props.initialWorkbook}
      workbookStore={props.store}
      workbookStorePersistent
      actor={props.actor}
      adapter={assistant}
      externalSourceAdapter={sources}
      onPersistenceStateChange={props.onSaveState}
    />
  </div>;
}
