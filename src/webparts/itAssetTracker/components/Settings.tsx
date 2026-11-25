import * as React from 'react';
import { DefaultButton, DetailsList, Dropdown, MessageBar, MessageBarType, Pivot, PivotItem, PrimaryButton, SelectionMode, Stack, Text, TextField } from '@fluentui/react';
import { IScreenProps } from '../services/models';
import { DataService } from '../services/DataService';
import { ASSET_TYPES, DEFAULT_REMINDER_DAYS, ENTITY_TYPES, LOCATIONS } from '../services/constants';
import { Badge, Busy, ErrorBar, PageHeader, col, opts, statusTone, useLoad } from './common';
import { errMsg, fmtDateTime, isEmail } from '../services/utils';

const Reminder: React.FC = () => {
  const { data, loading, error, reload } = useLoad(() => DataService.getSettings(), []);
  const [rec, setRec] = React.useState<string | undefined>();
  const [days, setDays] = React.useState<string | undefined>();
  const [msg, setMsg] = React.useState<{ t: string; ok: boolean } | undefined>();
  if (loading && !data) { return <Busy />; }
  const s = data || {};
  const domain = (DataService.me.EMail || '').split('@')[1];
  const recV = rec !== undefined ? rec : s.ReminderRecipients || '';
  const dayV = days !== undefined ? days : s.ReminderDays || DEFAULT_REMINDER_DAYS;

  const save = async (): Promise<void> => {
    const emails = recV.split(/[;,\s]+/).filter(x => !!x);
    if (!emails.length || emails.some(e => !isEmail(e) || (domain && e.split('@')[1].toLowerCase() !== domain.toLowerCase()))) {
      setMsg({ t: `Enter one or more valid organisation addresses (@${domain}), separated by semicolons.`, ok: false }); return;
    }
    const nums = dayV.split(',').map(x => x.trim());
    if (nums.some(n => !/^\d+$/.test(n) || Number(n) < 1)) { setMsg({ t: 'Thresholds must be positive whole numbers separated by commas, e.g. 30,15,7.', ok: false }); return; }
    try {
      await DataService.saveSetting('ReminderRecipients', emails.join(';'));
      await DataService.saveSetting('ReminderDays', nums.join(','));
      await DataService.saveSetting('TimeZone', s.TimeZone || 'Europe/London');
      setMsg({ t: 'Settings saved.', ok: true }); setRec(undefined); setDays(undefined); reload();
    } catch (e) { setMsg({ t: errMsg(e), ok: false }); }
  };
  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ maxWidth: 640, marginTop: 12 }}>
      <ErrorBar error={error} />
      {msg && <MessageBar messageBarType={msg.ok ? MessageBarType.success : MessageBarType.error} onDismiss={() => setMsg(undefined)}>{msg.t}</MessageBar>}
      <TextField label="Reminder recipients" multiline rows={2} value={recV} onChange={(_, v) => setRec(v || '')} description="Owner / IT admin only. Separate with semicolons." />
      <TextField label="Reminder thresholds (days)" value={dayV} onChange={(_, v) => setDays(v || '')} description="One alert per milestone per expiry date. Default 30,15,7." />
      <Text variant="small">The scheduled email flow (Power Automate F01, 08:00 Europe/London) reads these values from the App_Settings list.</Text>
      <PrimaryButton text="Save settings" onClick={save} styles={{ root: { width: 140 } }} />
      <Text variant="large" style={{ fontWeight: 600 }}>Governed reference values (changed only by a controlled release)</Text>
      <Text><b>Asset types:</b> {ASSET_TYPES.join(', ')}</Text>
      <Text><b>Locations:</b> {LOCATIONS.join(', ')}</Text>
    </Stack>
  );
};

const Audit: React.FC = () => {
  const [et, setEt] = React.useState('');
  const [oc, setOc] = React.useState('');
  const { data, loading, error } = useLoad(() => {
    const f: string[] = [];
    if (et) { f.push(`EntityType eq '${et}'`); }
    if (oc) { f.push(`Outcome eq '${oc}'`); }
    return DataService.getLog(200, f.join(' and ') || undefined);
  }, [et, oc]);
  return (
    <Stack tokens={{ childrenGap: 10 }} style={{ marginTop: 12 }}>
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <Dropdown label="Entity" selectedKey={et} options={opts(ENTITY_TYPES, 'All')} onChange={(_, o) => setEt(String(o!.key))} styles={{ root: { width: 170 } }} />
        <Dropdown label="Result" selectedKey={oc} options={opts(['Success', 'Failure'], 'All')} onChange={(_, o) => setOc(String(o!.key))} styles={{ root: { width: 120 } }} />
      </Stack>
      <ErrorBar error={error} />
      {loading ? <Busy /> : <DetailsList items={data || []} selectionMode={SelectionMode.none} compact columns={[
        col('ActionTime', 'Timestamp', 120, i => fmtDateTime(i.ActionTime)), col('EntityType', 'Entity', 100), col('EntityID', 'Item ID', 55), col('ActionName', 'Action', 80),
        col('PerformedBy', 'Actor', 110, i => (i.PerformedBy ? i.PerformedBy.Title : '—')), col('AfterSummary', 'Summary', 260, i => i.AfterSummary || i.Title),
        col('Outcome', 'Result', 70, i => <Badge text={i.Outcome} tone={statusTone(i.Outcome)} />), col('CorrelationID', 'Correlation', 80)
      ]} />}
      <Text variant="small">Audit history is read-only here; the app provides no delete action.</Text>
    </Stack>
  );
};

const Health: React.FC = () => {
  const rec = useLoad(() => DataService.reconcile(), []);
  const fail = useLoad(() => DataService.getFailedReminders(), []);
  return (
    <Stack tokens={{ childrenGap: 10 }} style={{ marginTop: 12 }}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}><Text variant="large" style={{ fontWeight: 600 }}>Reconciliation</Text><DefaultButton text="Re-run" onClick={rec.reload} /></Stack>
      <ErrorBar error={rec.error} />
      {rec.loading ? <Busy /> : rec.data && rec.data.length
        ? <MessageBar messageBarType={MessageBarType.severeWarning} isMultiline>{rec.data.map((x, i) => <div key={i}>• {x}</div>)}</MessageBar>
        : <MessageBar messageBarType={MessageBarType.success}>No inconsistencies found between asset status, open assignments and seat allocations.</MessageBar>}
      <Text variant="large" style={{ fontWeight: 600 }}>Reminder flow failures</Text>
      <ErrorBar error={fail.error} />
      {fail.loading ? <Busy /> : <DetailsList items={fail.data || []} selectionMode={SelectionMode.none} compact columns={[
        col('Title', 'Key', 220), col('EntityType', 'Type', 70), col('ExpiryDate', 'Expiry', 100, (i: any) => (i.ExpiryDate || '').substr(0, 10)), col('ThresholdDays', 'Days', 50), col('ErrorMessage', 'Error', 260)
      ]} />}
    </Stack>
  );
};

export const Settings: React.FC<IScreenProps> = ({ role }) => {
  if (role !== 'admin') { return <ErrorBar error="Access Denied: settings and audit are restricted to the Owner / Administrator (V12)." />; }
  return (
    <Stack>
      <PageHeader title="Settings and audit" />
      <Pivot aria-label="Settings sections">
        <PivotItem headerText="Reminders & reference data"><Reminder /></PivotItem>
        <PivotItem headerText="Audit log"><Audit /></PivotItem>
        <PivotItem headerText="Reconciliation & failures"><Health /></PivotItem>
      </Pivot>
    </Stack>
  );
};
