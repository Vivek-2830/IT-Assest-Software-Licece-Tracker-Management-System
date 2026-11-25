import * as React from 'react';
import { DefaultButton, DetailsList, MessageBarType, PrimaryButton, SelectionMode, Stack, Text } from '@fluentui/react';
import { IScreenProps } from '../services/models';
import { DataService } from '../services/DataService';
import { Badge, Busy, ErrorBar, ExpiryBadge, Kv, PageHeader, col, statusTone, useLoad } from './common';
import { DocumentsPanel } from './DocumentsPanel';
import { errMsg, fmt, fmtDateTime, money } from '../services/utils';
import { DEFAULT_CURRENCY } from '../services/constants';

export const AssetDetail: React.FC<IScreenProps> = ({ role, view, go }) => {
  const id = view.id!;
  const { data, loading, error, reload } = useLoad(async () => {
    const [asset, history, log] = await Promise.all([
      DataService.getAsset(id), DataService.getAssetAssignments(id), DataService.getLog(15, `EntityType eq 'Asset' and EntityID eq ${id}`)
    ]);
    return { asset, history, log };
  }, [id]);
  const [err, setErr] = React.useState<string | undefined>();
  const canEdit = role === 'admin' || role === 'editor';

  if (loading && !data) { return <Busy />; }
  if (!data) { return <ErrorBar error={error} />; }
  const { asset: a, history, log } = data;
  const open = history.filter(h => h.AssignmentStatus === 'Open')[0];

  const status = async (target: string, confirmText?: string): Promise<void> => {
    if (confirmText && !window.confirm(confirmText)) { return; }
    setErr(undefined);
    try { await DataService.changeAssetStatus(id, target, role === 'admin'); reload(); } catch (e) { setErr(errMsg(e)); reload(); }
  };

  return (
    <Stack tokens={{ childrenGap: 14 }}>
      <PageHeader title={`${a.AssetTag} – ${a.Title}`} subtitle={a.AssetType}>
        <Badge text={a.AssetStatus} tone={statusTone(a.AssetStatus)} />
      </PageHeader>
      <ErrorBar error={err || error} type={MessageBarType.error} onDismiss={() => setErr(undefined)} />
      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        <DefaultButton text="Back to list" iconProps={{ iconName: 'Back' }} onClick={() => go({ name: 'assets' })} />
        {canEdit && a.AssetStatus !== 'Retired' && <DefaultButton text="Edit" iconProps={{ iconName: 'Edit' }} onClick={() => go({ name: 'assetForm', id })} />}
        {canEdit && a.AssetStatus === 'Available' && <PrimaryButton text="Assign" onClick={() => go({ name: 'assignReturn', id })} />}
        {canEdit && a.AssetStatus === 'Assigned' && <PrimaryButton text="Return" onClick={() => go({ name: 'assignReturn', id })} />}
        {canEdit && a.AssetStatus === 'Available' && <DefaultButton text="Mark Repair" onClick={() => status('Repair')} />}
        {canEdit && a.AssetStatus === 'Repair' && <DefaultButton text="Mark Available" onClick={() => status('Available')} />}
        {canEdit && (a.AssetStatus === 'Available' || a.AssetStatus === 'Repair') &&
          <DefaultButton text="Retire" onClick={() => status('Retired', `Retire ${a.AssetTag}? Retired records become read-only until an Owner restores them.`)} />}
        {role === 'admin' && a.AssetStatus === 'Retired' && <DefaultButton text="Restore (Owner)" onClick={() => status('Available', `Restore ${a.AssetTag} to Available?`)} />}
      </Stack>

      <Text variant="large" style={{ fontWeight: 600 }}>Details</Text>
      <Stack horizontal wrap>
        <Kv label="Manufacturer" value={a.Manufacturer} /><Kv label="Model" value={a.Model} /><Kv label="Serial number" value={a.SerialNumber} />
        <Kv label="Purchase date" value={fmt(a.PurchaseDate)} /><Kv label="Purchase cost" value={a.PurchaseCost === undefined ? '—' : money(a.PurchaseCost, DEFAULT_CURRENCY)} />
        <Kv label="Supplier" value={a.Supplier} /><Kv label="Invoice reference" value={a.InvoiceReference} />
        <Kv label="Warranty end" value={<ExpiryBadge date={a.WarrantyEnd} />} /><Kv label="Location" value={a.OfficeLocation} />
        {a.ArchivedOn && <Kv label="Retired on" value={fmt(a.ArchivedOn)} />}
      </Stack>
      <Kv label="Notes" value={a.Notes} />

      <Text variant="large" style={{ fontWeight: 600 }}>Current holder</Text>
      {a.AssetStatus === 'Assigned' && open
        ? <Text>{open.AssignedTo ? open.AssignedTo.Title : a.CurrentHolder ? a.CurrentHolder.Title : 'Unknown'} since {fmt(open.AssignedDate)}</Text>
        : <Text>{a.AssetStatus === 'Assigned' ? 'Status is Assigned but no open assignment was found – run reconciliation (Settings).' : 'Available / not assigned'}</Text>}

      <Text variant="large" style={{ fontWeight: 600 }}>Assignment history</Text>
      <DetailsList items={history} selectionMode={SelectionMode.none} compact columns={[
        col('AssignedTo', 'Employee', 130, i => (i.AssignedTo ? i.AssignedTo.Title : '—')), col('AssignedDate', 'Assigned on', 90, i => fmt(i.AssignedDate)),
        col('ReturnedDate', 'Returned on', 90, i => fmt(i.ReturnedDate)), col('ConditionOut', 'Condition out', 80), col('ConditionIn', 'Condition in', 80),
        col('ReceivedBy', 'Returned to', 110, i => (i.ReceivedBy ? i.ReceivedBy.Title : '—')), col('Notes', 'Notes', 140)
      ]} />

      <Text variant="large" style={{ fontWeight: 600 }}>Documents</Text>
      <DocumentsPanel type="Asset" id={id} canUpload={canEdit && a.AssetStatus !== 'Retired'} />

      <Text variant="large" style={{ fontWeight: 600 }}>Change log</Text>
      <DetailsList items={log} selectionMode={SelectionMode.none} compact columns={[
        col('ActionTime', 'When', 120, i => fmtDateTime(i.ActionTime)), col('ActionName', 'Action', 80),
        col('AfterSummary', 'Summary', 260), col('PerformedBy', 'By', 110, i => (i.PerformedBy ? i.PerformedBy.Title : '—')),
        col('Outcome', 'Result', 70, i => <Badge text={i.Outcome} tone={statusTone(i.Outcome)} />)
      ]} />
    </Stack>
  );
};
