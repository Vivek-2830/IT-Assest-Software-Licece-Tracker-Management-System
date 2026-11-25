import * as React from 'react';
import { DefaultButton, Dropdown, MessageBarType, PrimaryButton, Stack, Text, TextField } from '@fluentui/react';
import { IPerson, IScreenProps } from '../services/models';
import { DataService } from '../services/DataService';
import { CONDITIONS } from '../services/constants';
import { Busy, DateField, ErrorBar, Kv, PageHeader, opts, useLoad } from './common';
import { PeoplePicker } from './PeoplePicker';
import { errMsg, fmt, sameOrAfter, today } from '../services/utils';

export const AssignReturn: React.FC<IScreenProps> = ({ view, go }) => {
  const id = view.id!;
  const { data, loading, error } = useLoad(async () => {
    const asset = await DataService.getAsset(id);
    const open = (await DataService.getAssetAssignments(id)).filter(h => h.AssignmentStatus === 'Open')[0];
    return { asset, open };
  }, [id]);
  const [person, setPerson] = React.useState<IPerson | undefined>();
  const [date, setDate] = React.useState<Date | undefined>(today());
  const [cond, setCond] = React.useState('Good');
  const [next, setNext] = React.useState('Available');
  const [comments, setComments] = React.useState('');
  const [errs, setErrs] = React.useState<{ [k: string]: string }>({});
  const [err, setErr] = React.useState<string | undefined>();
  const [busy, setBusy] = React.useState(false);

  if (loading && !data) { return <Busy />; }
  if (!data) { return <ErrorBar error={error} />; }
  const { asset, open } = data;
  const op = asset.AssetStatus === 'Available' ? 'Assign' : asset.AssetStatus === 'Assigned' ? 'Return' : '';
  if (!op) { return <ErrorBar error={`Asset ${asset.AssetTag} is ${asset.AssetStatus}; it cannot be assigned or returned (V06).`} />; }

  const confirm = async (): Promise<void> => {
    const e: { [k: string]: string } = {};
    if (op === 'Assign' && !person) { e.person = 'Select the employee (V01).'; }
    if (!date) { e.date = 'Required (V01).'; }
    if (op === 'Return' && date && open && open.AssignedDate && !sameOrAfter(date, open.AssignedDate)) { e.date = `Return date cannot be before ${fmt(open.AssignedDate)} (V08).`; }
    setErrs(e);
    if (Object.keys(e).length) { return; }
    setBusy(true); setErr(undefined);
    try {
      if (op === 'Assign') { await DataService.assignAsset(id, person!, date!, cond, comments); }
      else { await DataService.returnAsset(id, date!, cond, next, comments); }
      go({ name: 'assetDetail', id });
    } catch (x) { setErr(errMsg(x)); setBusy(false); }
  };

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ maxWidth: 640 }}>
      <PageHeader title={`${op} asset`} subtitle="Records chain of custody. State is re-checked immediately before writing." />
      <ErrorBar error={err} type={MessageBarType.error} />
      <Kv label="Asset" value={`${asset.AssetTag} – ${asset.Title}`} />
      {op === 'Return' && <Kv label="Currently held by" value={`${open && open.AssignedTo ? open.AssignedTo.Title : '—'} (since ${open ? fmt(open.AssignedDate) : '—'})`} />}
      {op === 'Assign' && <PeoplePicker label="Assigned to" required value={person} onChange={setPerson} error={errs.person} />}
      <DateField label={op === 'Assign' ? 'Assigned date' : 'Return date'} value={date} onChange={setDate} required error={errs.date} />
      <Dropdown label={op === 'Assign' ? 'Condition out' : 'Condition in'} selectedKey={cond} options={opts(CONDITIONS)} onChange={(_, o) => { setCond(String(o!.key)); if (op === 'Return') { setNext(o!.key === 'Damaged' ? 'Repair' : 'Available'); } }} />
      {op === 'Return' && <Dropdown label="Set asset status to" selectedKey={next} options={opts(['Available', 'Repair'])} onChange={(_, o) => setNext(String(o!.key))} />}
      {op === 'Return' && <Text variant="small">Received by: {DataService.me.Title} (you)</Text>}
      <TextField label="Comments" multiline rows={3} value={comments} onChange={(_, v) => setComments(v || '')} />
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text={busy ? 'Working…' : `Confirm ${op.toLowerCase()}`} disabled={busy} onClick={confirm} />
        <DefaultButton text="Cancel" onClick={() => go({ name: 'assetDetail', id })} />
      </Stack>
    </Stack>
  );
};
