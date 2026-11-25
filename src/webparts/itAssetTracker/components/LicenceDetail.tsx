import * as React from 'react';
import { DefaultButton, DetailsList, MessageBarType, PrimaryButton, SelectionMode, Stack, Text } from '@fluentui/react';
import { IScreenProps } from '../services/models';
import { DataService, isSeatManaged } from '../services/DataService';
import { Badge, Busy, ErrorBar, ExpiryBadge, Kv, LinkCell, PageHeader, col, statusTone, useLoad } from './common';
import { DocumentsPanel } from './DocumentsPanel';
import { RenewDialog } from './RenewDialog';
import { errMsg, fmt, money } from '../services/utils';

export const LicenceDetail: React.FC<IScreenProps> = ({ role, view, go }) => {
  const id = view.id!;
  const { data, loading, error, reload } = useLoad(async () => {
    const [lic, allocs, assets] = await Promise.all([DataService.getLicence(id), DataService.getAllocations(id), DataService.getAssets()]);
    return { lic, allocs, assets };
  }, [id]);
  const [err, setErr] = React.useState<string | undefined>();
  const [renew, setRenew] = React.useState(false);
  const canEdit = role === 'admin' || role === 'editor';

  if (loading && !data) { return <Busy />; }
  if (!data) { return <ErrorBar error={error} />; }
  const { lic: l, allocs, assets } = data;
  const open = allocs.filter(a => a.AllocationStatus === 'Open');
  const history = allocs.filter(a => a.AllocationStatus !== 'Open');
  const seatMode = isSeatManaged(l);
  const tagOf = (aid?: number): string => { const a = assets.filter(x => x.Id === aid)[0]; return a ? `${a.AssetTag} – ${a.Title}` : `Asset #${aid}`; };
  const target = (a: any): string => (a.TargetPerson ? a.TargetPerson.Title : a.TargetAssetID ? tagOf(a.TargetAssetID) : '—');

  const cancel = async (): Promise<void> => {
    const msg = open.length
      ? `${open.length} seat(s) are still allocated (V11). Cancelling blocks NEW allocations but keeps the existing open allocations as history. Continue?`
      : `Cancel ${l.LicenceCode}? No new allocations will be possible.`;
    if (!window.confirm(msg)) { return; }
    try { await DataService.cancelLicence(id); reload(); } catch (e) { setErr(errMsg(e)); }
  };

  return (
    <Stack tokens={{ childrenGap: 14 }}>
      <PageHeader title={`${l.LicenceCode} – ${l.Title}`} subtitle={`${l.Vendor} • ${l.LicenceModel}`}><Badge text={l.LicenceStatus} tone={statusTone(l.LicenceStatus)} /></PageHeader>
      <ErrorBar error={err || error} type={MessageBarType.error} onDismiss={() => setErr(undefined)} />
      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        <DefaultButton text="Back to list" iconProps={{ iconName: 'Back' }} onClick={() => go({ name: 'licences' })} />
        {canEdit && l.LicenceStatus !== 'Archived' && <DefaultButton text="Edit" iconProps={{ iconName: 'Edit' }} onClick={() => go({ name: 'licenceForm', id })} />}
        {canEdit && seatMode && l.LicenceStatus === 'Active' && <PrimaryButton text="Allocate seat" onClick={() => go({ name: 'seatAllocation', id, params: { mode: 'allocate' } })} />}
        {canEdit && l.LicenceStatus !== 'Cancelled' && l.LicenceStatus !== 'Archived' && <DefaultButton text="Renew / extend" onClick={() => setRenew(true)} />}
        {canEdit && l.LicenceStatus === 'Active' && <DefaultButton text="Mark Cancelled" onClick={cancel} />}
      </Stack>
      {renew && <RenewDialog licence={l} onClose={c => { setRenew(false); if (c) { reload(); } }} />}

      <Text variant="large" style={{ fontWeight: 600 }}>Summary</Text>
      <Stack horizontal wrap>
        <Kv label="Term" value={`${fmt(l.StartDate)} → ${fmt(l.RenewalDate)}`} /><Kv label="Renewal" value={<ExpiryBadge date={l.RenewalDate} />} />
        <Kv label="Renewable / auto-renew" value={`${l.Renewable ? 'Yes' : 'No'} / ${l.AutoRenew ? 'Yes' : 'No'}`} />
        <Kv label="Amount" value={`${money(l.Amount, l.CurrencyCode)} (${l.BillingFrequency})`} /><Kv label="Last renewed" value={fmt(l.LastRenewedOn)} />
        <Kv label="Vendor contact" value={l.VendorContact} /><Kv label="Contract reference" value={l.ContractReference} />
      </Stack>
      <Kv label="Notes" value={l.Notes} />

      <Text variant="large" style={{ fontWeight: 600 }}>Seats</Text>
      {seatMode ? (
        <Stack horizontal wrap>
          <Kv label="Purchased" value={l.PurchasedSeats} /><Kv label="Assigned" value={open.length} />
          <Kv label="Available" value={(l.PurchasedSeats || 0) - open.length} />
        </Stack>
      ) : <Text>N/A – this licence model has no seat allocation.</Text>}

      {seatMode && <>
        <Text variant="large" style={{ fontWeight: 600 }}>Current allocations</Text>
        <DetailsList items={open} selectionMode={SelectionMode.none} compact columns={[
          col('t', 'Employee / device', 200, target), col('AssignedDate', 'Allocated', 100, a => fmt(a.AssignedDate)), col('AllocationStatus', 'Status', 70),
          col('act', '', 80, a => (canEdit ? <LinkCell text="Reclaim" onClick={() => go({ name: 'seatAllocation', id, params: { mode: 'reclaim', allocationId: a.Id } })} /> : null))
        ]} />
        <Text variant="large" style={{ fontWeight: 600 }}>History</Text>
        <DetailsList items={history} selectionMode={SelectionMode.none} compact columns={[
          col('t', 'Employee / device', 200, target), col('AssignedDate', 'Allocated', 100, a => fmt(a.AssignedDate)), col('ReclaimedDate', 'Reclaimed', 100, a => fmt(a.ReclaimedDate)), col('Notes', 'Notes', 160)
        ]} />
      </>}

      <Text variant="large" style={{ fontWeight: 600 }}>Documents</Text>
      <DocumentsPanel type="Licence" id={id} canUpload={canEdit} />
    </Stack>
  );
};
