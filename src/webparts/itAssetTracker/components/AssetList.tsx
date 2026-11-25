import * as React from 'react';
import { DefaultButton, DetailsList, Dropdown, PrimaryButton, SelectionMode, Stack, TextField, Toggle } from '@fluentui/react';
import { IAsset, IScreenProps } from '../services/models';
import { DataService } from '../services/DataService';
import { ASSET_STATUSES, ASSET_TYPES, LOCATIONS } from '../services/constants';
import { Badge, Busy, ErrorBar, ExpiryBadge, LinkCell, PageHeader, col, opts, statusTone, useLoad } from './common';

export const AssetList: React.FC<IScreenProps> = ({ role, go, view }) => {
  const p = view.params || {};
  const { data, loading, error } = useLoad(() => DataService.getAssets(), []);
  const [q, setQ] = React.useState('');
  const [status, setStatus] = React.useState<string>(p.status || '');
  const [type, setType] = React.useState('');
  const [holder, setHolder] = React.useState('');
  const [loc, setLoc] = React.useState('');
  const [hideRetired, setHide] = React.useState<boolean>(!!p.hideRetired);
  const [sort, setSort] = React.useState<{ key: string; desc: boolean }>({ key: 'AssetTag', desc: false });
  const canEdit = role === 'admin' || role === 'editor';

  if (loading && !data) { return <Busy />; }
  const all: IAsset[] = data || [];
  const holders = Array.from(new Set(all.filter(a => a.CurrentHolder).map(a => a.CurrentHolder!.Title))).sort();
  const sw = (s?: string): boolean => !!s && s.toLowerCase().indexOf(q.trim().toLowerCase()) === 0;   // starts-with per spec

  let rows = all.filter(a =>
    (!q.trim() || sw(a.AssetTag) || sw(a.Title) || sw(a.SerialNumber) || sw(a.Model)) &&
    (!status || a.AssetStatus === status) && (!type || a.AssetType === type) && (!loc || a.OfficeLocation === loc) &&
    (!holder || (a.CurrentHolder && a.CurrentHolder.Title === holder)) && (!hideRetired || a.AssetStatus !== 'Retired'));
  rows = rows.sort((a, b) => {
    const av = sort.key === 'WarrantyEnd' ? (a.WarrantyEnd ? a.WarrantyEnd.getTime() : 9e15) : a.AssetTag;
    const bv = sort.key === 'WarrantyEnd' ? (b.WarrantyEnd ? b.WarrantyEnd.getTime() : 9e15) : b.AssetTag;
    return (av < bv ? -1 : av > bv ? 1 : 0) * (sort.desc ? -1 : 1);
  });
  const clear = (): void => { setQ(''); setStatus(''); setType(''); setHolder(''); setLoc(''); setHide(false); };
  const sortable = (key: string) => ({ isSorted: sort.key === key, isSortedDescending: sort.desc, onColumnClick: () => setSort({ key, desc: sort.key === key ? !sort.desc : false }) });

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <PageHeader title="Assets" subtitle={`${rows.length} of ${all.length} shown`}>
        {canEdit && <PrimaryButton text="Add Asset" iconProps={{ iconName: 'Add' }} onClick={() => go({ name: 'assetForm' })} />}
      </PageHeader>
      <ErrorBar error={error} />
      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 10 }}>
        <TextField label="Search (tag, name, serial, model – starts with)" value={q} onChange={(_, v) => setQ(v || '')} styles={{ root: { width: 300 } }} iconProps={{ iconName: 'Search' }} />
        <Dropdown label="Status" selectedKey={status} options={opts(ASSET_STATUSES, 'All')} onChange={(_, o) => setStatus(String(o!.key))} styles={{ root: { width: 130 } }} />
        <Dropdown label="Type" selectedKey={type} options={opts(ASSET_TYPES, 'All')} onChange={(_, o) => setType(String(o!.key))} styles={{ root: { width: 130 } }} />
        <Dropdown label="Assigned to" selectedKey={holder} options={opts(holders, 'Anyone')} onChange={(_, o) => setHolder(String(o!.key))} styles={{ root: { width: 170 } }} />
        <Dropdown label="Location" selectedKey={loc} options={opts(LOCATIONS, 'All')} onChange={(_, o) => setLoc(String(o!.key))} styles={{ root: { width: 130 } }} />
        <Toggle label="Hide retired" checked={hideRetired} onChange={(_, c) => setHide(!!c)} />
        <DefaultButton text="Clear filters" onClick={clear} />
      </Stack>
      <DetailsList
        items={rows} selectionMode={SelectionMode.none} onItemInvoked={(i: IAsset) => go({ name: 'assetDetail', id: i.Id })}
        columns={[
          col('AssetTag', 'Asset tag', 80, (i: IAsset) => <LinkCell text={i.AssetTag} onClick={() => go({ name: 'assetDetail', id: i.Id })} />, sortable('AssetTag')),
          col('Title', 'Asset name', 160), col('AssetType', 'Category', 80), col('Model', 'Model', 100),
          col('CurrentHolder', 'Custodian', 120, (i: IAsset) => (i.CurrentHolder ? i.CurrentHolder.Title : '—')),
          col('AssetStatus', 'Status', 80, (i: IAsset) => <Badge text={i.AssetStatus} tone={statusTone(i.AssetStatus)} />),
          col('WarrantyEnd', 'Warranty end', 170, (i: IAsset) => <ExpiryBadge date={i.WarrantyEnd} />, sortable('WarrantyEnd')),
          col('view', '', 60, (i: IAsset) => <LinkCell text="View" onClick={() => go({ name: 'assetDetail', id: i.Id })} />)
        ]}
      />
    </Stack>
  );
};
