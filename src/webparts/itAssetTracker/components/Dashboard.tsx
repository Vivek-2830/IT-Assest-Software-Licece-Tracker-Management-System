import * as React from 'react';
import { DefaultButton, DetailsList, Link, PrimaryButton, SelectionMode, Stack, Text } from '@fluentui/react';
import { IScreenProps } from '../services/models';
import { DataService, isSeatManaged } from '../services/DataService';
import { Badge, Busy, ErrorBar, ExpiryBadge, PageHeader, col, statusTone, useLoad } from './common';
import { daysTo, fmtDateTime, money } from '../services/utils';

const Tile: React.FC<{ label: string; value: number | string; onClick?: () => void; tone?: string }> = ({ label, value, onClick, tone }) => (
  <button
    onClick={onClick} disabled={!onClick} aria-label={`${label}: ${value}`}
    style={{ flex: '1 1 150px', minWidth: 150, textAlign: 'left', padding: 12, border: '1px solid #c8c6c4', borderLeft: `4px solid ${tone || '#0078d4'}`, background: '#fff', cursor: onClick ? 'pointer' : 'default' }}>
    <div style={{ fontSize: 12, color: '#605e5c' }}>{label}</div>
    <div style={{ fontSize: 28, fontWeight: 600 }}>{value}</div>
  </button>
);

export const Dashboard: React.FC<IScreenProps> = ({ role, go }) => {
  const { data, loading, error, reload } = useLoad(async () => {
    const [assets, licences, allocs, log] = await Promise.all([DataService.getAssets(), DataService.getLicences(), DataService.getOpenAllocations(), DataService.getLog(10)]);
    return { assets, licences, allocs, log };
  }, []);
  const canEdit = role === 'admin' || role === 'editor';
  if (loading && !data) { return <Busy />; }
  if (!data) { return <ErrorBar error={error} />; }
  const { assets, licences, allocs, log } = data;

  const live = assets.filter(a => a.AssetStatus !== 'Retired');
  const count = (s: string): number => assets.filter(a => a.AssetStatus === s).length;
  const activeLic = licences.filter(l => l.LicenceStatus === 'Active');
  const seatLic = activeLic.filter(isSeatManaged);
  const purchased = seatLic.reduce((n, l) => n + (l.PurchasedSeats || 0), 0);
  const assigned = allocs.filter(a => seatLic.some(l => l.Id === a.LicenceID)).length;
  const renewals = activeLic.filter(l => l.RenewalDate && daysTo(l.RenewalDate) <= 30);
  const warranties = live.filter(a => a.WarrantyEnd && daysTo(a.WarrantyEnd) <= 30).sort((a, b) => a.WarrantyEnd!.getTime() - b.WarrantyEnd!.getTime());

  // Optional estimated GBP run-rate: monthly x12, quarterly x4, yearly x1. Other currencies / one-time / other are excluded and labelled.
  const mult: { [k: string]: number } = { Monthly: 12, Quarterly: 4, Yearly: 1 };
  let run = 0; let excluded = 0;
  activeLic.forEach(l => {
    if (l.Amount !== undefined && l.CurrencyCode === 'GBP' && mult[l.BillingFrequency]) { run += l.Amount * mult[l.BillingFrequency]; } else if (l.Amount) { excluded++; }
  });

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <PageHeader title="Dashboard" subtitle="Counts include current records only; retired assets are excluded from in-use KPIs.">
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={reload} />
        {canEdit && <PrimaryButton text="Add Asset" iconProps={{ iconName: 'Add' }} onClick={() => go({ name: 'assetForm' })} />}
        {canEdit && <DefaultButton text="Add Licence" iconProps={{ iconName: 'Add' }} onClick={() => go({ name: 'licenceForm' })} />}
        {canEdit && <DefaultButton text="Assign Asset" onClick={() => go({ name: 'assets', params: { status: 'Available' } })} />}
        <DefaultButton text="View Renewals" onClick={() => go({ name: 'renewals' })} />
      </PageHeader>
      <ErrorBar error={error} />

      <Text variant="large" style={{ fontWeight: 600 }}>Assets</Text>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <Tile label="Total active assets" value={live.length} tone="#0078d4" onClick={() => go({ name: 'assets', params: { hideRetired: true } })} />
        <Tile label="Assigned" value={count('Assigned')} tone="#0f548c" onClick={() => go({ name: 'assets', params: { status: 'Assigned' } })} />
        <Tile label="Available" value={count('Available')} tone="#107c10" onClick={() => go({ name: 'assets', params: { status: 'Available' } })} />
        <Tile label="Under repair" value={count('Repair')} tone="#c19c00" onClick={() => go({ name: 'assets', params: { status: 'Repair' } })} />
        <Tile label="Retired" value={count('Retired')} tone="#8a8886" onClick={() => go({ name: 'assets', params: { status: 'Retired' } })} />
      </Stack>

      <Text variant="large" style={{ fontWeight: 600 }}>Software licences</Text>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <Tile label="Active subscriptions" value={activeLic.length} onClick={() => go({ name: 'licences', params: { status: 'Active' } })} />
        <Tile label="Seats purchased" value={purchased} />
        <Tile label="Seats assigned" value={assigned} />
        <Tile label="Renewals due ≤ 30 days" value={renewals.length} tone="#c19c00" onClick={() => go({ name: 'licences', params: { status: 'Active', renewal: 30 } })} />
        <Tile label="Est. annual run-rate (GBP)" value={money(run, '').trim()} />
      </Stack>
      {excluded > 0 && <Text variant="small">Run-rate counts GBP monthly/quarterly/yearly items only; {excluded} item(s) in other currencies, one-time or other billing are excluded.</Text>}

      <Stack horizontal wrap tokens={{ childrenGap: 24 }}>
        <Stack style={{ flex: '1 1 380px' }} tokens={{ childrenGap: 6 }}>
          <Text variant="large" style={{ fontWeight: 600 }}>Warranties expiring ≤ 30 days (incl. overdue)</Text>
          {warranties.length === 0 ? <Text>None.</Text> : warranties.map(a => (
            <Stack key={a.Id} horizontal horizontalAlign="space-between"><Link onClick={() => go({ name: 'assetDetail', id: a.Id })}>{a.AssetTag} – {a.Title}</Link><ExpiryBadge date={a.WarrantyEnd} /></Stack>
          ))}
        </Stack>
        <Stack style={{ flex: '1 1 380px' }} tokens={{ childrenGap: 6 }}>
          <Text variant="large" style={{ fontWeight: 600 }}>Subscriptions expiring ≤ 30 days (incl. overdue)</Text>
          {renewals.length === 0 ? <Text>None.</Text> : renewals.sort((a, b) => a.RenewalDate!.getTime() - b.RenewalDate!.getTime()).map(l => (
            <Stack key={l.Id} horizontal horizontalAlign="space-between"><Link onClick={() => go({ name: 'licenceDetail', id: l.Id })}>{l.LicenceCode} – {l.Title}</Link><ExpiryBadge date={l.RenewalDate} /></Stack>
          ))}
        </Stack>
      </Stack>

      <Text variant="large" style={{ fontWeight: 600 }}>Recent activity</Text>
      <DetailsList
        items={log} selectionMode={SelectionMode.none} compact
        columns={[
          col('ActionTime', 'Timestamp', 130, i => fmtDateTime(i.ActionTime)),
          col('EntityType', 'Entity', 100), col('ActionName', 'Action', 90),
          col('AfterSummary', 'Summary', 220, i => i.AfterSummary || i.Title),
          col('PerformedBy', 'User', 110, i => (i.PerformedBy ? i.PerformedBy.Title : '—')),
          col('Outcome', 'Result', 70, i => <Badge text={i.Outcome} tone={statusTone(i.Outcome)} />)
        ]}
      />
    </Stack>
  );
};
