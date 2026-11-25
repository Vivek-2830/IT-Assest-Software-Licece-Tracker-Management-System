import * as React from 'react';
import { DetailsList, Dropdown, Pivot, PivotItem, SelectionMode, Stack } from '@fluentui/react';
import { IAsset, ILicence, IScreenProps } from '../services/models';
import { DataService } from '../services/DataService';
import { Badge, Busy, ErrorBar, ExpiryBadge, LinkCell, PageHeader, col, statusTone, useLoad } from './common';
import { RenewDialog } from './RenewDialog';
import { daysTo, fmt } from '../services/utils';

const dayText = (d: Date | undefined): string => {
  if (!d) { return '—'; }
  const n = daysTo(d);
  return n < 0 ? `Overdue by ${-n}` : String(n);
};

export const Renewals: React.FC<IScreenProps> = ({ role, go }) => {
  const { data, loading, error, reload } = useLoad(async () => {
    const [licences, assets] = await Promise.all([DataService.getLicences(), DataService.getAssets()]);
    return { licences, assets };
  }, []);
  const [horizon, setHorizon] = React.useState(30);
  const [renewing, setRenewing] = React.useState<ILicence | undefined>();
  const canEdit = role === 'admin' || role === 'editor';
  if (loading && !data) { return <Busy />; }
  const lics = (data ? data.licences : []).filter(l => l.LicenceStatus === 'Active' && l.RenewalDate && daysTo(l.RenewalDate) <= horizon)
    .sort((a, b) => a.RenewalDate!.getTime() - b.RenewalDate!.getTime());
  const ws = (data ? data.assets : []).filter(a => a.AssetStatus !== 'Retired' && a.WarrantyEnd && daysTo(a.WarrantyEnd) <= horizon)
    .sort((a, b) => a.WarrantyEnd!.getTime() - b.WarrantyEnd!.getTime());   // overdue items stay in the list

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <PageHeader title="Renewals and warranties" subtitle="Expired items stay listed as overdue until renewed or retired." />
      <ErrorBar error={error} />
      <Dropdown label="Time horizon" selectedKey={horizon} styles={{ root: { width: 160 } }}
        options={[15, 30, 60, 90].map(d => ({ key: d, text: `${d} days` }))} onChange={(_, o) => setHorizon(Number(o!.key))} />
      {renewing && <RenewDialog licence={renewing} onClose={c => { setRenewing(undefined); if (c) { reload(); } }} />}
      <Pivot aria-label="Renewals and warranties">
        <PivotItem headerText={`Software renewals (${lics.length})`}>
          <DetailsList items={lics} selectionMode={SelectionMode.none} columns={[
            col('Title', 'Name', 180, (l: ILicence) => <LinkCell text={`${l.LicenceCode} – ${l.Title}`} onClick={() => go({ name: 'licenceDetail', id: l.Id })} />),
            col('RenewalDate', 'Expiry', 170, (l: ILicence) => <ExpiryBadge date={l.RenewalDate} />), col('days', 'Days remaining', 90, (l: ILicence) => dayText(l.RenewalDate)),
            col('VendorContact', 'Owner / contact', 130, (l: ILicence) => l.VendorContact || l.Vendor),
            col('LicenceStatus', 'Status', 80, (l: ILicence) => <Badge text={l.LicenceStatus} tone={statusTone(l.LicenceStatus)} />),
            col('act', '', 90, (l: ILicence) => (canEdit ? <LinkCell text="Renew…" onClick={() => setRenewing(l)} /> : null))
          ]} />
        </PivotItem>
        <PivotItem headerText={`Warranty expiries (${ws.length})`}>
          <DetailsList items={ws} selectionMode={SelectionMode.none} columns={[
            col('Title', 'Name', 180, (a: IAsset) => <LinkCell text={`${a.AssetTag} – ${a.Title}`} onClick={() => go({ name: 'assetDetail', id: a.Id })} />),
            col('WarrantyEnd', 'Expiry', 170, (a: IAsset) => <ExpiryBadge date={a.WarrantyEnd} />), col('days', 'Days remaining', 90, (a: IAsset) => dayText(a.WarrantyEnd)),
            col('holder', 'Assigned to', 130, (a: IAsset) => (a.CurrentHolder ? a.CurrentHolder.Title : '—')),
            col('AssetStatus', 'Status', 80, (a: IAsset) => <Badge text={a.AssetStatus} tone={statusTone(a.AssetStatus)} />),
            col('act', '', 90, (a: IAsset) => (canEdit ? <LinkCell text="Edit date" onClick={() => go({ name: 'assetForm', id: a.Id })} /> : null))
          ]} />
        </PivotItem>
      </Pivot>
    </Stack>
  );
};
