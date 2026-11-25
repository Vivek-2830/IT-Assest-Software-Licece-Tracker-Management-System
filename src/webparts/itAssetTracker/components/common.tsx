import * as React from 'react';
import {
  DatePicker, DayOfWeek, IColumn, IDropdownOption, Link, MessageBar, MessageBarType, Spinner, SpinnerSize, Stack, Text
} from '@fluentui/react';
import { errMsg, fmt, daysTo } from '../services/utils';

export interface ILoad<T> { data?: T; loading: boolean; error?: string; reload: () => void; }

export function useLoad<T>(fn: () => Promise<T>, deps: any[] = []): ILoad<T> {
  const [state, setState] = React.useState<{ data?: T; loading: boolean; error?: string }>({ loading: true });
  const [tick, setTick] = React.useState(0);
  React.useEffect(() => {
    let off = false;
    setState(s => ({ data: s.data, loading: true }));
    fn().then(d => { if (!off) { setState({ data: d, loading: false }); } })
      .catch(e => { if (!off) { setState({ loading: false, error: errMsg(e) }); } });
    return () => { off = true; };
  }, [...deps, tick]);
  return { ...state, reload: () => setTick(t => t + 1) };
}

export const opts = (a: string[], blank?: string): IDropdownOption[] => {
  const o = a.map(x => ({ key: x, text: x }));
  return blank !== undefined ? [{ key: '', text: blank }, ...o] : o;
};

export const Busy: React.FC = () => <Spinner size={SpinnerSize.large} label="Loading…" style={{ padding: 40 }} />;

export const ErrorBar: React.FC<{ error?: string; onDismiss?: () => void; type?: MessageBarType }> = ({ error, onDismiss, type }) =>
  error ? <MessageBar messageBarType={type || MessageBarType.error} isMultiline onDismiss={onDismiss} role="alert">{error}</MessageBar> : null;

export const PageHeader: React.FC<{ title: string; subtitle?: string }> = ({ title, subtitle, children }) => (
  <Stack horizontal verticalAlign="center" horizontalAlign="space-between" wrap tokens={{ childrenGap: 12 }} style={{ margin: '4px 0 16px' }}>
    <Stack>
      <Text variant="xLarge" block style={{ fontWeight: 600 }}>{title}</Text>
      {subtitle && <Text variant="small" style={{ color: '#605e5c' }}>{subtitle}</Text>}
    </Stack>
    <Stack horizontal wrap tokens={{ childrenGap: 8 }}>{children}</Stack>
  </Stack>
);

export const Kv: React.FC<{ label: string; value?: React.ReactNode }> = ({ label, value }) => (
  <div style={{ minWidth: 200, flex: '1 1 200px', padding: '4px 0' }}>
    <Text variant="small" block style={{ color: '#605e5c' }}>{label}</Text>
    <Text variant="medium" block>{value === undefined || value === null || value === '' ? '—' : value}</Text>
  </div>
);

const TONES: { [k: string]: [string, string] } = {
  green: ['#dff6dd', '#107c10'], blue: ['#deecf9', '#0f548c'], amber: ['#fff4ce', '#8a6100'],
  grey: ['#edebe9', '#323130'], red: ['#fde7e9', '#a4262c']
};
export const Badge: React.FC<{ text: string; tone?: string }> = ({ text, tone }) => {
  const t = TONES[tone || 'grey'];
  return <span style={{ background: t[0], color: t[1], padding: '2px 8px', borderRadius: 10, fontSize: 12, fontWeight: 600, whiteSpace: 'nowrap' }}>{text}</span>;
};

export const statusTone = (s: string): string =>
  ({ Available: 'green', Assigned: 'blue', Repair: 'amber', Retired: 'grey', Active: 'green', Cancelled: 'grey', Expired: 'red', Archived: 'grey',
    Pending: 'amber', Approved: 'blue', Rejected: 'red', Revoked: 'grey', Failure: 'red', Success: 'green' } as any)[s] || 'grey';

/** Expiry badge – text + colour (never colour alone). */
export const ExpiryBadge: React.FC<{ date?: Date }> = ({ date }) => {
  if (!date) { return <span>—</span>; }
  const d = daysTo(date);
  const tone = d < 0 ? 'red' : d <= 7 ? 'red' : d <= 30 ? 'amber' : 'green';
  const label = d < 0 ? `Overdue by ${-d}d` : d === 0 ? 'Due today' : `${d}d left`;
  return <span><span style={{ marginRight: 6 }}>{fmt(date)}</span><Badge text={label} tone={tone} /></span>;
};

export const col = (key: string, name: string, minWidth: number, onRender?: (i: any) => React.ReactNode, extra?: Partial<IColumn>): IColumn =>
  ({ key, name, fieldName: key, minWidth, maxWidth: minWidth * 2.2, isResizable: true, onRender, ...(extra || {}) });

export const LinkCell: React.FC<{ text: string; onClick: () => void }> = ({ text, onClick }) => <Link onClick={onClick}>{text}</Link>;

/** Date-only field with calendar + clear link. */
export const DateField: React.FC<{
  label: string; value?: Date; onChange: (d?: Date) => void; required?: boolean; error?: string; minDate?: Date; disabled?: boolean;
}> = ({ label, value, onChange, required, error, minDate, disabled }) => (
  <div>
    <DatePicker
      label={label} isRequired={required} value={value} onSelectDate={d => onChange(d || undefined)}
      formatDate={d => (d ? fmt(d) : '')} firstDayOfWeek={DayOfWeek.Monday} minDate={minDate} disabled={disabled}
      placeholder="dd MMM yyyy" ariaLabel={label}
      textField={{ errorMessage: error }}
    />
    {value && !required && !disabled && <Link onClick={() => onChange(undefined)} style={{ fontSize: 12 }}>Clear {label}</Link>}
  </div>
);
