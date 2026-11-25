import * as React from 'react';
import { Checkbox, DefaultButton, DetailsList, Dropdown, MessageBar, MessageBarType, PrimaryButton, SelectionMode, Stack, Text, TextField } from '@fluentui/react';
import { IAccessRequest, IScreenProps } from '../services/models';
import { DataService } from '../services/DataService';
import { ACCESS_ROLES } from '../services/constants';
import { Badge, Busy, ErrorBar, PageHeader, col, opts, statusTone, useLoad } from './common';
import { errMsg, fmtDateTime, isEmail } from '../services/utils';

export const AccessManagement: React.FC<IScreenProps> = ({ role }) => {
  const { data, loading, error, reload } = useLoad(() => DataService.getAccessRequests(), []);
  const [email, setEmail] = React.useState('');
  const [rl, setRl] = React.useState('Editor');
  const [notes, setNotes] = React.useState('');
  const [ack, setAck] = React.useState(false);
  const [err, setErr] = React.useState<string | undefined>();
  const [busy, setBusy] = React.useState(false);

  if (role !== 'admin') { return <ErrorBar error="Access Denied: only the Owner / Administrator can manage access requests (V12)." />; }
  if (loading && !data) { return <Busy />; }
  const all = data || [];
  const domain = (DataService.me.EMail || '').split('@')[1];

  const run = async (fn: () => Promise<void>): Promise<void> => {
    setBusy(true); setErr(undefined);
    try { await fn(); reload(); } catch (e) { setErr(errMsg(e)); }
    setBusy(false);
  };
  const add = (): Promise<void> => {
    const e = email.trim().toLowerCase();
    if (!isEmail(e)) { setErr('Enter a valid work email.'); return Promise.resolve(); }
    if (domain && e.split('@')[1] !== domain.toLowerCase()) { setErr(`Only organisational addresses (@${domain}) are eligible.`); return Promise.resolve(); }
    return run(async () => { await DataService.addAccessRequest(e, rl, notes); setEmail(''); setNotes(''); });
  };
  const act = (r: IAccessRequest, status: string, confirmText?: string): void => {
    if (confirmText && !window.confirm(confirmText)) { return; }
    run(() => DataService.setAccessStatus(r, status, all)).catch(() => undefined);
  };

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <PageHeader title="User access management" subtitle="Tracks requests and approvals only." />
      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        This list does <b>not</b> grant or revoke access. A trusted Microsoft 365 administrator must add/remove the user in the Entra groups ITAM-Admins / ITAM-Editors / ITAM-Viewers
        and apply app/site sharing. Revoke the real access first, then mark Revoked. Mark Active only after testing actual access (sign out / in).
      </MessageBar>
      <ErrorBar error={err || error} type={MessageBarType.error} onDismiss={() => setErr(undefined)} />
      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 10 }}>
        <TextField label="User work email" required value={email} onChange={(_, v) => setEmail(v || '')} styles={{ root: { width: 280 } }} />
        <Dropdown label="Role" required selectedKey={rl} options={opts(ACCESS_ROLES)} onChange={(_, o) => setRl(String(o!.key))} styles={{ root: { width: 120 } }} />
        <TextField label="Notes" value={notes} onChange={(_, v) => setNotes(v || '')} styles={{ root: { width: 280 } }} />
        <PrimaryButton text="Create request" disabled={busy} onClick={add} />
      </Stack>
      <Checkbox label="I confirm group membership and app access were actually applied/removed in Microsoft 365 before I use Mark Active / Revoke" checked={ack} onChange={(_, c) => setAck(!!c)} />
      <DetailsList items={all} selectionMode={SelectionMode.none} columns={[
        col('UserEmail', 'Email', 190), col('RequestedRole', 'Role', 70),
        col('AccessStatus', 'Status', 80, (r: IAccessRequest) => <Badge text={r.AccessStatus} tone={statusTone(r.AccessStatus)} />),
        col('RequestedBy', 'Requested', 150, (r: IAccessRequest) => `${r.RequestedBy ? r.RequestedBy.Title : '—'} ${fmtDateTime(r.RequestedOn)}`),
        col('ApprovedBy', 'Decision', 150, (r: IAccessRequest) => (r.ApprovedBy ? `${r.ApprovedBy.Title} ${fmtDateTime(r.ApprovedOn)}` : '—')),
        col('ProvisionedBy', 'Provisioned', 150, (r: IAccessRequest) => (r.ProvisionedBy ? `${r.ProvisionedBy.Title} ${fmtDateTime(r.ProvisionedOn)}` : '—')),
        col('Notes', 'Notes', 120),
        col('act', 'Actions', 280, (r: IAccessRequest) => (
          <Stack horizontal tokens={{ childrenGap: 6 }}>
            {r.AccessStatus === 'Pending' && <DefaultButton text="Approve" disabled={busy} onClick={() => act(r, 'Approved')} />}
            {r.AccessStatus === 'Pending' && <DefaultButton text="Reject" disabled={busy} onClick={() => act(r, 'Rejected')} />}
            {r.AccessStatus === 'Approved' && <DefaultButton text="Mark Active" disabled={busy || !ack} onClick={() => act(r, 'Active')} />}
            {(r.AccessStatus === 'Active' || r.AccessStatus === 'Approved') && <DefaultButton text="Revoke" disabled={busy || !ack} onClick={() => act(r, 'Revoked', `Have you already removed ${r.UserEmail} from the groups/app/site? Mark as Revoked?`)} />}
          </Stack>)) 
      ]} />
      <Text variant="small">Initial Owner and IT Admin accounts should be recorded here as Active admin / editor entries so the last-Owner safeguard (V15) works.</Text>
    </Stack>
  );
};
