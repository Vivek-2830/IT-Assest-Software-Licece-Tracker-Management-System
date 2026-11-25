import * as React from 'react';
import { DefaultButton, Dropdown, MessageBarType, PrimaryButton, Stack, Text, TextField } from '@fluentui/react';
import { IPerson, IScreenProps } from '../services/models';
import { DataService, isSeatManaged } from '../services/DataService';
import { Busy, DateField, ErrorBar, Kv, PageHeader, useLoad } from './common';
import { PeoplePicker } from './PeoplePicker';
import { errMsg, today } from '../services/utils';

export const SeatAllocation: React.FC<IScreenProps> = ({ view, go }) => {
  const id = view.id!;
  const mode: string = (view.params && view.params.mode) || 'allocate';
  const allocationId: number | undefined = view.params && view.params.allocationId;
  const { data, loading, error } = useLoad(async () => {
    const [lic, allocs, assets] = await Promise.all([DataService.getLicence(id), DataService.getAllocations(id), DataService.getAssets()]);
    return { lic, allocs, assets };
  }, [id]);
  const [person, setPerson] = React.useState<IPerson | undefined>();
  const [assetId, setAssetId] = React.useState<number | undefined>();
  const [date, setDate] = React.useState<Date | undefined>(today());
  const [notes, setNotes] = React.useState('');
  const [errs, setErrs] = React.useState<{ [k: string]: string }>({});
  const [err, setErr] = React.useState<string | undefined>();
  const [busy, setBusy] = React.useState(false);

  if (loading && !data) { return <Busy />; }
  if (!data) { return <ErrorBar error={error} />; }
  const { lic, allocs, assets } = data;
  if (!isSeatManaged(lic)) { return <ErrorBar error="This licence model does not support seat allocation." />; }
  const open = allocs.filter(a => a.AllocationStatus === 'Open');
  const remaining = (lic.PurchasedSeats || 0) - open.length;
  const deviceMode = lic.LicenceModel === 'Device seat';
  const back = (): void => go({ name: 'licenceDetail', id });
  const reclaimTarget = open.filter(a => a.Id === allocationId)[0];

  const submit = async (): Promise<void> => {
    const e: { [k: string]: string } = {};
    if (!date) { e.date = 'Required (V01).'; }
    if (mode === 'allocate') {
      if (deviceMode && !assetId) { e.target = 'Select the device (V01).'; }
      if (!deviceMode && !person) { e.target = 'Select the employee (V01).'; }
    }
    setErrs(e);
    if (Object.keys(e).length) { return; }
    setBusy(true); setErr(undefined);
    try {
      if (mode === 'reclaim') { await DataService.reclaimSeat(allocationId!, date!); }
      else {
        const asset = assets.filter(a => a.Id === assetId)[0];
        await DataService.allocateSeat(id, deviceMode ? { assetId, assetTag: asset ? asset.AssetTag : undefined } : { person }, date!, notes);
      }
      back();
    } catch (x) { setErr(errMsg(x)); setBusy(false); }
  };

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ maxWidth: 640 }}>
      <PageHeader title={mode === 'reclaim' ? 'Reclaim seat' : 'Allocate seat'} subtitle="Seat limit and duplicate rules are re-checked before and after the write." />
      <ErrorBar error={err} type={MessageBarType.error} />
      <Stack horizontal wrap><Kv label="Licence" value={`${lic.LicenceCode} – ${lic.Title}`} /><Kv label="Seats remaining" value={`${remaining} of ${lic.PurchasedSeats}`} /></Stack>
      {mode === 'allocate' && (deviceMode
        ? <Dropdown label="Device" required placeholder="Select device" selectedKey={assetId} errorMessage={errs.target}
          options={assets.filter(a => a.AssetStatus !== 'Retired').sort((a, b) => a.AssetTag.localeCompare(b.AssetTag)).map(a => ({ key: a.Id, text: `${a.AssetTag} – ${a.Title}` }))}
          onChange={(_, o) => setAssetId(Number(o!.key))} />
        : <PeoplePicker label="Employee" required value={person} onChange={setPerson} error={errs.target} />)}
      {mode === 'reclaim' && <Text>Reclaiming the seat held by <b>{reclaimTarget ? (reclaimTarget.TargetPerson ? reclaimTarget.TargetPerson.Title : `asset #${reclaimTarget.TargetAssetID}`) : 'unknown'}</b>.</Text>}
      <DateField label={mode === 'reclaim' ? 'Reclaimed date' : 'Allocated date'} value={date} onChange={setDate} required error={errs.date} />
      {mode === 'allocate' && <TextField label="Notes" multiline rows={2} value={notes} onChange={(_, v) => setNotes(v || '')} />}
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text={busy ? 'Working…' : mode === 'reclaim' ? 'Reclaim' : 'Allocate'} disabled={busy || (mode === 'allocate' && remaining <= 0)} onClick={submit} />
        <DefaultButton text="Cancel" onClick={back} />
      </Stack>
      {mode === 'allocate' && remaining <= 0 && <Text style={{ color: '#a4262c' }}>No seats remaining (V10).</Text>}
    </Stack>
  );
};
