import * as React from 'react';
import { DefaultButton, Dropdown, MessageBarType, PrimaryButton, Stack, TextField } from '@fluentui/react';
import { IAsset, IAssetInput, IScreenProps } from '../services/models';
import { DataService } from '../services/DataService';
import { ASSET_TYPES, DEFAULT_CURRENCY, LOCATIONS } from '../services/constants';
import { Busy, DateField, ErrorBar, PageHeader, opts, useLoad } from './common';
import { errMsg, sameOrAfter, today } from '../services/utils';

interface IForm { Title: string; AssetTag: string; AssetType: string; Manufacturer: string; Model: string; SerialNumber: string;
  PurchaseDate?: Date; PurchaseCost: string; Supplier: string; InvoiceReference: string; WarrantyEnd?: Date; OfficeLocation: string; Notes: string; }

const blank: IForm = { Title: '', AssetTag: '', AssetType: '', Manufacturer: '', Model: '', SerialNumber: '', PurchaseCost: '', Supplier: '',
  InvoiceReference: '', OfficeLocation: 'Unspecified', Notes: '' };

export const AssetForm: React.FC<IScreenProps> = ({ view, go }) => {
  const editing = !!view.id;
  const existing = useLoad<IAsset | undefined>(async () => (view.id ? DataService.getAsset(view.id) : undefined), [view.id]);
  const [f, setF] = React.useState<IForm>(blank);
  const [errs, setErrs] = React.useState<{ [k: string]: string }>({});
  const [err, setErr] = React.useState<string | undefined>();
  const [busy, setBusy] = React.useState(false);
  const loadedModified = React.useRef('');

  React.useEffect(() => {
    const a = existing.data;
    if (a) {
      loadedModified.current = a.Modified;
      setF({ Title: a.Title, AssetTag: a.AssetTag, AssetType: a.AssetType, Manufacturer: a.Manufacturer || '', Model: a.Model || '', SerialNumber: a.SerialNumber || '',
        PurchaseDate: a.PurchaseDate, PurchaseCost: a.PurchaseCost === undefined ? '' : String(a.PurchaseCost), Supplier: a.Supplier || '',
        InvoiceReference: a.InvoiceReference || '', WarrantyEnd: a.WarrantyEnd, OfficeLocation: a.OfficeLocation, Notes: a.Notes || '' });
    }
  }, [existing.data]);

  if (editing && existing.loading && !existing.data) { return <Busy />; }
  if (existing.data && existing.data.AssetStatus === 'Retired') {
    return <ErrorBar error="Retired records are read-only. An Owner can restore the asset from its detail page." />;
  }
  const set = (k: keyof IForm, v: any): void => setF(s => ({ ...s, [k]: v }));

  const validate = (): boolean => {
    const e: { [k: string]: string } = {};
    if (!f.Title.trim()) { e.Title = 'Required (V01).'; }
    if (f.Title.length > 150) { e.Title = 'Maximum 150 characters.'; }
    if (!f.AssetType) { e.AssetType = 'Required (V01).'; }
    if (!f.OfficeLocation) { e.OfficeLocation = 'Required (V01).'; }
    if (editing && !f.AssetTag.trim()) { e.AssetTag = 'Required (V01).'; }
    if (f.Manufacturer.length > 100 || f.Model.length > 100 || f.SerialNumber.length > 100) { e.Model = 'Manufacturer, model and serial are limited to 100 characters.'; }
    if (f.PurchaseCost !== '' && (isNaN(Number(f.PurchaseCost)) || Number(f.PurchaseCost) < 0)) { e.PurchaseCost = 'Cost must be a number ≥ 0 (V05).'; }
    if (f.PurchaseDate && f.PurchaseDate > today()) { e.PurchaseDate = 'Purchase date cannot be in the future.'; }
    if (f.PurchaseDate && f.WarrantyEnd && !sameOrAfter(f.WarrantyEnd, f.PurchaseDate)) { e.WarrantyEnd = 'Warranty end cannot be before the purchase date (V04).'; }
    setErrs(e);
    return Object.keys(e).length === 0;
  };

  const save = async (): Promise<void> => {
    setErr(undefined);
    if (!validate()) { return; }
    const input: IAssetInput = {
      Title: f.Title.trim(), AssetTag: f.AssetTag.trim().toUpperCase(), AssetType: f.AssetType, Manufacturer: f.Manufacturer.trim(), Model: f.Model.trim(),
      SerialNumber: f.SerialNumber.trim(), PurchaseDate: f.PurchaseDate, PurchaseCost: f.PurchaseCost === '' ? undefined : Number(f.PurchaseCost),
      Supplier: f.Supplier.trim(), InvoiceReference: f.InvoiceReference.trim(), WarrantyEnd: f.WarrantyEnd, OfficeLocation: f.OfficeLocation, Notes: f.Notes
    };
    setBusy(true);
    try {
      let id = view.id!;
      if (editing) { await DataService.updateAsset(id, input, loadedModified.current); } else { id = await DataService.createAsset(input); }
      go({ name: 'assetDetail', id });
    } catch (e) { setErr(errMsg(e)); setBusy(false); }   // form data is preserved; never shows false success (V13)
  };

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ maxWidth: 860 }}>
      <PageHeader title={editing ? `Edit asset ${f.AssetTag}` : 'Add asset'} subtitle="Status and custodian change only through Assign / Return / status actions." />
      <ErrorBar error={err} type={MessageBarType.error} />
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <TextField label="Asset tag" required={editing} value={f.AssetTag} onChange={(_, v) => set('AssetTag', v || '')} errorMessage={errs.AssetTag}
          placeholder="Leave blank to auto-generate (IT-0001)" description="Unique" styles={{ root: { width: 260 } }} />
        <TextField label="Title" required value={f.Title} onChange={(_, v) => set('Title', v || '')} errorMessage={errs.Title} styles={{ root: { width: 380 } }} />
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <Dropdown label="Asset type" required selectedKey={f.AssetType || undefined} options={opts(ASSET_TYPES)} placeholder="Select" onChange={(_, o) => set('AssetType', String(o!.key))} errorMessage={errs.AssetType} styles={{ root: { width: 200 } }} />
        <Dropdown label="Office location" required selectedKey={f.OfficeLocation} options={opts(LOCATIONS)} onChange={(_, o) => set('OfficeLocation', String(o!.key))} errorMessage={errs.OfficeLocation} styles={{ root: { width: 200 } }} />
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <TextField label="Manufacturer" value={f.Manufacturer} onChange={(_, v) => set('Manufacturer', v || '')} styles={{ root: { width: 250 } }} />
        <TextField label="Model" value={f.Model} onChange={(_, v) => set('Model', v || '')} errorMessage={errs.Model} styles={{ root: { width: 250 } }} />
        <TextField label="Serial number" value={f.SerialNumber} onChange={(_, v) => set('SerialNumber', v || '')} description="Unique when provided" styles={{ root: { width: 250 } }} />
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <DateField label="Purchase date" value={f.PurchaseDate} onChange={d => set('PurchaseDate', d)} error={errs.PurchaseDate} />
        <TextField label={`Purchase cost (${DEFAULT_CURRENCY})`} type="number" min={0} value={f.PurchaseCost} onChange={(_, v) => set('PurchaseCost', v || '')} errorMessage={errs.PurchaseCost} styles={{ root: { width: 180 } }} />
        <DateField label="Warranty end" value={f.WarrantyEnd} onChange={d => set('WarrantyEnd', d)} error={errs.WarrantyEnd} />
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <TextField label="Supplier" value={f.Supplier} onChange={(_, v) => set('Supplier', v || '')} styles={{ root: { width: 250 } }} />
        <TextField label="Invoice reference" value={f.InvoiceReference} onChange={(_, v) => set('InvoiceReference', v || '')} styles={{ root: { width: 250 } }} />
      </Stack>
      <TextField label="Notes" multiline rows={4} value={f.Notes} onChange={(_, v) => set('Notes', v || '')} description="Do not record passwords or keys." />
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text={busy ? 'Saving…' : 'Save'} disabled={busy} onClick={save} />
        <DefaultButton text="Cancel" onClick={() => go(editing ? { name: 'assetDetail', id: view.id } : { name: 'assets' })} />
      </Stack>
    </Stack>
  );
};
