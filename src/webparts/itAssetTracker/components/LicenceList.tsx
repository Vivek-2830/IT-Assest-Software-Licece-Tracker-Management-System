import * as React from 'react';
import { DefaultButton, DetailsList, Dropdown, PrimaryButton, SelectionMode, Stack, TextField } from '@fluentui/react';
import { ILicence, IScreenProps } from '../services/models';
import { DataService, isSeatManaged } from '../services/DataService';
import { BILLING, LICENCE_MODELS, LICENCE_STATUSES } from '../services/constants';
import { Badge, Busy, ErrorBar, ExpiryBadge, LinkCell, PageHeader, col, opts, statusTone, useLoad } from './common';
import { daysTo, money } from '../services/utils';

export const LicenceList: React.FC<IScreenProps> = ({ role, go, view }) => {
  const p = view.params || {};
  const { data, loading, error } = useLoad(async () => {
    const [licences, allocs] = await Promise.all([DataService.getLicences(), DataService.getOpenAllocations()]);
    return { licences, allocs };
  }, []);
  const [q, setQ] = React.useState('');
  const [status, setStatus] = React.useState<string>(p.status || '');
  const [model, setModel] = React.useState('');
  const [renewal, setRenewal] = React.useState<string>(p.renewal ? String(p.renewal) : '');
  const canEdit = role === 'admin' || role === 'editor';

  if (loading && !data) { return <Busy />; }
  const licences = data ? data.licences : [];
  const seats = (l: ILicence): number => (data ? data.allocs.filter(a => a.LicenceID === l.Id).length : 0);
  const sw = (s?: string): boolean => !!s && s.toLowerCase().indexOf(q.trim().toLowerCase()) === 0;
  const rows = licences.filter(l =>
    (!q.trim() || sw(l.Title) || sw(l.Vendor) || sw(l.LicenceCode)) && (!status || l.LicenceStatus === status) && (!model || l.LicenceModel === model) &&
    (!renewal || (l.RenewalDate && daysTo(l.RenewalDate) <= Number(renewal)))).sort((a, b) => a.LicenceCode.localeCompare(b.LicenceCode));

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <PageHeader title="Software licences" subtitle={`${rows.length} of ${licences.length} shown`}>
        {canEdit && <PrimaryButton text="Add Licence" iconProps={{ iconName: 'Add' }} onClick={() => go({ name: 'licenceForm' })} />}
      </PageHeader>
      <ErrorBar error={error} />
      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 10 }}>
        <TextField label="Search (software, vendor – starts with)" value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { width: 280 } }} iconProps={{ iconName: 'Search' }} />
        <Dropdown label="Status" selectedKey={status} options={opts(LICENCE_STATUSES, 'All')} onChange={(_, o) => setStatus(String(o!.key))} styles={{ root: { width: 130 } }} />
        <Dropdown label="Licence model" selectedKey={model} options={opts(LICENCE_MODELS, 'All')} onChange={(_, o) => setModel(String(o!.key))} styles={{ root: { width: 170 } }} />
        <Dropdown label="Renewal" selectedKey={renewal} options={[{ key: '', text: 'Any' }, { key: '30', text: 'Within 30 days' }, { key: '60', text: 'Within 60 days' }]} onChange={(_, o) => setRenewal(String(o!.key))} styles={{ root: { width: 150 } }} />
        <DefaultButton text="Clear" onClick={() => { setQ(''); setStatus(''); setModel(''); setRenewal(''); }} />
      </Stack>
      <DetailsList items={rows} selectionMode={SelectionMode.none} onItemInvoked={(l: ILicence) => go({ name: 'licenceDetail', id: l.Id })} columns={[
        col('LicenceCode', 'Code', 70, (l: ILicence) => <LinkCell text={l.LicenceCode} onClick={() => go({ name: 'licenceDetail', id: l.Id })} />),
        col('Title', 'Software', 140), col('Vendor', 'Vendor', 110), col('LicenceModel', 'Model', 110),
        col('PurchasedSeats', 'Purchased', 70, (l: ILicence) => (isSeatManaged(l) ? l.PurchasedSeats : 'N/A')),
        col('assigned', 'Assigned', 70, (l: ILicence) => (isSeatManaged(l) ? seats(l) : 'N/A')),
        col('RenewalDate', 'Renewal', 170, (l: ILicence) => <ExpiryBadge date={l.RenewalDate} />),
        col('Amount', 'Cost', 130, (l: ILicence) => `${money(l.Amount, l.CurrencyCode)} / ${l.BillingFrequency}`),
        col('LicenceStatus', 'Status', 80, (l: ILicence) => <Badge text={l.LicenceStatus} tone={statusTone(l.LicenceStatus)} />)
      ]} />
    </Stack>
  );
};
