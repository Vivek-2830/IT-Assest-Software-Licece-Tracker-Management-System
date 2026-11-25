import * as React from 'react';
import { DefaultButton, Dropdown, MessageBar, MessageBarType, PrimaryButton, Stack, TextField, Toggle } from '@fluentui/react';
import { ILicence, ILicenceInput, IScreenProps } from '../services/models';
import { DataService } from '../services/DataService';
import { BILLING, CURRENCIES, DEFAULT_CURRENCY, LICENCE_MODELS, LICENCE_STATUSES, SEAT_MODELS } from '../services/constants';
import { Busy, DateField, ErrorBar, PageHeader, opts, useLoad } from './common';
import { errMsg, sameOrAfter } from '../services/utils';

interface IForm { LicenceCode: string; Title: string; Vendor: string; LicenceModel: string; PurchasedSeats: string; StartDate?: Date; RenewalDate?: Date;
  Renewable: boolean; AutoRenew: boolean; BillingFrequency: string; Amount: string; CurrencyCode: string; LicenceStatus: string;
  VendorContact: string; ContractReference: string; Notes: string; }
const blank: IForm = { LicenceCode: '', Title: '', Vendor: '', LicenceModel: '', PurchasedSeats: '', Renewable: true, AutoRenew: false, BillingFrequency: 'Yearly',
  Amount: '', CurrencyCode: DEFAULT_CURRENCY, LicenceStatus: 'Active', VendorContact: '', ContractReference: '', Notes: '' };

export const LicenceForm: React.FC<IScreenProps> = ({ view, go }) => {
  const editing = !!view.id;
  const existing = useLoad<ILicence | undefined>(async () => (view.id ? DataService.getLicence(view.id) : undefined), [view.id]);
  const [f, setF] = React.useState<IForm>(blank);
  const [errs, setErrs] = React.useState<{ [k: string]: string }>({});
  const [err, setErr] = React.useState<string | undefined>();
  const [busy, setBusy] = React.useState(false);
  const mod = React.useRef('');

  React.useEffect(() => {
    const l = existing.data;
    if (l) {
      mod.current = l.Modified;
      setF({ LicenceCode: l.LicenceCode, Title: l.Title, Vendor: l.Vendor, LicenceModel: l.LicenceModel, PurchasedSeats: l.PurchasedSeats === undefined ? '' : String(l.PurchasedSeats),
        StartDate: l.StartDate, RenewalDate: l.RenewalDate, Renewable: l.Renewable, AutoRenew: l.AutoRenew, BillingFrequency: l.BillingFrequency,
        Amount: l.Amount === undefined ? '' : String(l.Amount), CurrencyCode: l.CurrencyCode, LicenceStatus: l.LicenceStatus,
        VendorContact: l.VendorContact || '', ContractReference: l.ContractReference || '', Notes: l.Notes || '' });
    }
  }, [existing.data]);

  if (editing && existing.loading && !existing.data) { return <Busy />; }
  const set = (k: keyof IForm, v: any): void => setF(s => ({ ...s, [k]: v }));
  const seatManaged = SEAT_MODELS.indexOf(f.LicenceModel) >= 0 || f.LicenceModel === 'Perpetual';

  const validate = (): boolean => {
    const e: { [k: string]: string } = {};
    if (!f.Title.trim()) { e.Title = 'Required (V01).'; }
    if (!f.Vendor.trim()) { e.Vendor = 'Required (V01).'; }
    if (f.Title.length > 150 || f.Vendor.length > 150) { e.Title = 'Maximum 150 characters.'; }
    if (!f.LicenceModel) { e.LicenceModel = 'Required (V01).'; }
    if (editing && !f.LicenceCode.trim()) { e.LicenceCode = 'Required (V01).'; }
    if (SEAT_MODELS.indexOf(f.LicenceModel) >= 0 && (f.PurchasedSeats === '' || !/^\d+$/.test(f.PurchasedSeats) || Number(f.PurchasedSeats) < 1)) { e.PurchasedSeats = 'Seat-based models need a positive whole number of seats (V05).'; }
    if (f.PurchasedSeats !== '' && (!/^\d+$/.test(f.PurchasedSeats) || Number(f.PurchasedSeats) < 1)) { e.PurchasedSeats = 'Seats must be a positive whole number (V05).'; }
    if (f.Amount !== '' && (isNaN(Number(f.Amount)) || Number(f.Amount) < 0)) { e.Amount = 'Amount must be ≥ 0 (V05).'; }
    if (f.Renewable && !f.RenewalDate) { e.RenewalDate = 'Required when Renewable = Yes (V01).'; }
    if (f.StartDate && f.RenewalDate && !sameOrAfter(f.RenewalDate, f.StartDate)) { e.RenewalDate = 'Renewal date cannot be before the start date (V04).'; }
    setErrs(e);
    return Object.keys(e).length === 0;
  };

  const save = async (): Promise<void> => {
    setErr(undefined);
    if (!validate()) { return; }
    const input: ILicenceInput = {
      LicenceCode: f.LicenceCode.trim().toUpperCase(), Title: f.Title.trim(), Vendor: f.Vendor.trim(), LicenceModel: f.LicenceModel,
      PurchasedSeats: f.PurchasedSeats === '' || !seatManaged ? undefined : Number(f.PurchasedSeats), StartDate: f.StartDate,
      RenewalDate: f.RenewalDate, Renewable: f.Renewable, AutoRenew: f.Renewable && f.AutoRenew, BillingFrequency: f.BillingFrequency,
      Amount: f.Amount === '' ? undefined : Number(f.Amount), CurrencyCode: f.CurrencyCode, LicenceStatus: f.LicenceStatus,
      VendorContact: f.VendorContact.trim(), ContractReference: f.ContractReference.trim(), Notes: f.Notes
    };
    setBusy(true);
    try {
      let id = view.id!;
      if (editing) { await DataService.updateLicence(id, input, mod.current); } else { id = await DataService.createLicence(input); }
      go({ name: 'licenceDetail', id });
    } catch (e) { setErr(errMsg(e)); setBusy(false); }
  };

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ maxWidth: 860 }}>
      <PageHeader title={editing ? `Edit licence ${f.LicenceCode}` : 'Add licence'} subtitle="One row per separately purchased subscription / contract." />
      <ErrorBar error={err} type={MessageBarType.error} />
      <MessageBar messageBarType={MessageBarType.warning}>Never store licence / activation keys or passwords in this record.</MessageBar>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <TextField label="Licence code" required={editing} value={f.LicenceCode} onChange={(_, v) => set('LicenceCode', v || '')} errorMessage={errs.LicenceCode} placeholder="Blank = auto (LIC-0001)" styles={{ root: { width: 220 } }} />
        <TextField label="Software name" required value={f.Title} onChange={(_, v) => set('Title', v || '')} errorMessage={errs.Title} styles={{ root: { width: 280 } }} />
        <TextField label="Vendor" required value={f.Vendor} onChange={(_, v) => set('Vendor', v || '')} errorMessage={errs.Vendor} styles={{ root: { width: 240 } }} />
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <Dropdown label="Licence model" required selectedKey={f.LicenceModel || undefined} placeholder="Select" options={opts(LICENCE_MODELS)} onChange={(_, o) => set('LicenceModel', String(o!.key))} errorMessage={errs.LicenceModel} styles={{ root: { width: 220 } }} />
        <TextField label="Purchased seats" type="number" min={1} disabled={!seatManaged} value={seatManaged ? f.PurchasedSeats : ''} onChange={(_, v) => set('PurchasedSeats', v || '')} errorMessage={errs.PurchasedSeats} description={seatManaged ? undefined : 'Not applicable for this model'} styles={{ root: { width: 160 } }} />
        {editing && <Dropdown label="Status" selectedKey={f.LicenceStatus} options={opts(LICENCE_STATUSES)} onChange={(_, o) => set('LicenceStatus', String(o!.key))} styles={{ root: { width: 140 } }} />}
      </Stack>
      <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 12 }}>
        <DateField label="Start date" value={f.StartDate} onChange={d => set('StartDate', d)} />
        <DateField label="Renewal date" value={f.RenewalDate} onChange={d => set('RenewalDate', d)} required={f.Renewable} error={errs.RenewalDate} />
        <Toggle label="Renewable" checked={f.Renewable} onChange={(_, c) => set('Renewable', !!c)} onText="Yes" offText="No" />
        <Toggle label="Auto-renew (vendor)" checked={f.Renewable && f.AutoRenew} disabled={!f.Renewable} onChange={(_, c) => set('AutoRenew', !!c)} onText="Yes" offText="No" />
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <Dropdown label="Billing frequency" required selectedKey={f.BillingFrequency} options={opts(BILLING)} onChange={(_, o) => set('BillingFrequency', String(o!.key))} styles={{ root: { width: 170 } }} />
        <TextField label="Amount" type="number" min={0} value={f.Amount} onChange={(_, v) => set('Amount', v || '')} errorMessage={errs.Amount} styles={{ root: { width: 160 } }} />
        <Dropdown label="Currency" required selectedKey={f.CurrencyCode} options={opts(CURRENCIES)} onChange={(_, o) => set('CurrencyCode', String(o!.key))} styles={{ root: { width: 110 } }} />
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <TextField label="Vendor contact" value={f.VendorContact} onChange={(_, v) => set('VendorContact', v || '')} styles={{ root: { width: 300 } }} />
        <TextField label="Contract reference" value={f.ContractReference} onChange={(_, v) => set('ContractReference', v || '')} styles={{ root: { width: 300 } }} />
      </Stack>
      <TextField label="Notes" multiline rows={4} value={f.Notes} onChange={(_, v) => set('Notes', v || '')} />
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text={busy ? 'Saving…' : 'Save'} disabled={busy} onClick={save} />
        <DefaultButton text="Cancel" onClick={() => go(editing ? { name: 'licenceDetail', id: view.id } : { name: 'licences' })} />
      </Stack>
    </Stack>
  );
};
