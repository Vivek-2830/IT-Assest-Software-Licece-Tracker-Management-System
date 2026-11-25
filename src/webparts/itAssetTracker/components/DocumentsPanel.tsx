import * as React from 'react';
import { DefaultButton, Dropdown, Link, MessageBarType, Stack, Text } from '@fluentui/react';
import { DataService } from '../services/DataService';
import { DOC_TYPES } from '../services/constants';
import { Busy, ErrorBar, opts, useLoad } from './common';
import { errMsg } from '../services/utils';

/** Links to receipts / warranties / contracts held in the restricted ITAM_Documents library. */
export const DocumentsPanel: React.FC<{ type: 'Asset' | 'Licence'; id: number; canUpload: boolean }> = ({ type, id, canUpload }) => {
  const docs = useLoad(() => DataService.getDocs(type, id), [type, id]);
  const [docType, setDocType] = React.useState('Invoice');
  const [busy, setBusy] = React.useState(false);
  const [err, setErr] = React.useState<string | undefined>();
  const fileRef = React.useRef<HTMLInputElement>(null);

  const upload = async (): Promise<void> => {
    const f = fileRef.current && fileRef.current.files && fileRef.current.files[0];
    if (!f) { setErr('Choose a file first.'); return; }
    setBusy(true); setErr(undefined);
    try { await DataService.uploadDoc(type, id, docType, f); if (fileRef.current) { fileRef.current.value = ''; } docs.reload(); }
    catch (e) { setErr(errMsg(e)); }
    setBusy(false);
  };

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <ErrorBar error={err || docs.error} type={MessageBarType.error} />
      {docs.loading ? <Busy /> : (docs.data && docs.data.length
        ? <ul style={{ margin: 0, paddingLeft: 18 }}>{docs.data.map(d => <li key={d.Id}><Link href={d.FileRef} target="_blank" rel="noopener noreferrer">{d.FileLeafRef}</Link> <Text variant="small">({d.DocumentType || 'Other'})</Text></li>)}</ul>
        : <Text>No documents linked.</Text>)}
      {canUpload && (
        <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
          <Dropdown label="Document type" selectedKey={docType} options={opts(DOC_TYPES)} onChange={(_, o) => setDocType(String(o!.key))} styles={{ root: { width: 160 } }} />
          <input type="file" ref={fileRef} aria-label="Choose document" />
          <DefaultButton text={busy ? 'Uploading…' : 'Upload'} disabled={busy} onClick={upload} iconProps={{ iconName: 'Upload' }} />
        </Stack>
      )}
    </Stack>
  );
};
