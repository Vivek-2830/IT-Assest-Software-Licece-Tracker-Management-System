import * as React from 'react';
import { IPersonaProps, Label, NormalPeoplePicker, Text } from '@fluentui/react';
import { IPerson } from '../services/models';
import { DataService } from '../services/DataService';
import { errMsg } from '../services/utils';

interface IProps { label: string; value?: IPerson; required?: boolean; disabled?: boolean; error?: string; onChange: (p?: IPerson) => void; }

/** Directory person picker. Picking a person only records a custodian; it never grants app access. */
export const PeoplePicker: React.FC<IProps> = ({ label, value, required, disabled, error, onChange }) => {
  const [err, setErr] = React.useState<string | undefined>();
  const selected: IPersonaProps[] = value ? [{ text: value.Title, secondaryText: value.EMail, id: String(value.Id) }] : [];

  const resolve = async (filter: string): Promise<IPersonaProps[]> => {
    if (!filter || filter.length < 2) { return []; }
    try {
      const r = await DataService.searchPeople(filter);
      return r.map(x => ({ text: x.DisplayText, secondaryText: x.Email, id: x.Key }));
    } catch (e) { setErr(errMsg(e)); return []; }
  };

  const changed = async (items?: IPersonaProps[]): Promise<void> => {
    if (!items || !items.length) { onChange(undefined); return; }
    const it = items[0];
    if (it.id && /^\d+$/.test(it.id)) { onChange(value); return; }
    try { onChange(await DataService.ensureUser(it.id as string)); setErr(undefined); }
    catch (e) { setErr(errMsg(e)); }
  };

  return (
    <div>
      <Label required={required}>{label}</Label>
      <NormalPeoplePicker
        onResolveSuggestions={resolve} onEmptyResolveSuggestions={() => []} selectedItems={selected} onChange={changed}
        itemLimit={1} disabled={disabled} resolveDelay={300} inputProps={{ placeholder: 'Type a name or email', 'aria-label': label }}
      />
      {(error || err) && <Text variant="small" style={{ color: '#a4262c' }}>{error || err}</Text>}
    </div>
  );
};
