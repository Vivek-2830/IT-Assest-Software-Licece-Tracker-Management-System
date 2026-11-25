import * as React from 'react';
import { Dialog, DialogFooter, DialogType, PrimaryButton, DefaultButton, MessageBarType } from '@fluentui/react';
import { ILicence } from '../services/models';
import { DataService } from '../services/DataService';
import { DateField, ErrorBar } from './common';
import { errMsg } from '../services/utils';

/** Extend a renewal. Only records the new date; the system never auto-renews or claims payment. */
export const RenewDialog: React.FC<{ licence: ILicence; onClose: (changed: boolean) => void }> = ({ licence, onClose }) => {
  const [date, setDate] = React.useState<Date | undefined>(licence.RenewalDate);
  const [err, setErr] = React.useState<string | undefined>();
  const [busy, setBusy] = React.useState(false);
  const save = async (): Promise<void> => {
    if (!date) { setErr('Choose the new renewal date.'); return; }
    setBusy(true);
    try { await DataService.renewLicence(licence.Id, date); onClose(true); } catch (e) { setErr(errMsg(e)); setBusy(false); }
  };
  return (
    <Dialog hidden={false} onDismiss={() => onClose(false)} dialogContentProps={{ type: DialogType.normal, title: `Extend renewal – ${licence.LicenceCode}`, subText: `${licence.Title}. Set the new renewal date after the renewal is confirmed.` }}>
      <ErrorBar error={err} type={MessageBarType.error} />
      <DateField label="New renewal date" value={date} onChange={setDate} required />
      <DialogFooter>
        <PrimaryButton text="Save" onClick={save} disabled={busy} />
        <DefaultButton text="Cancel" onClick={() => onClose(false)} />
      </DialogFooter>
    </Dialog>
  );
};
