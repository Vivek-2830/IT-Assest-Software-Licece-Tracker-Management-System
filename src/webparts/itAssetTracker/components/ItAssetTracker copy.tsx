import * as React from 'react';
import { initializeIcons, MessageBarType, Nav, Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';
import { IScreenProps, IView, Role, ViewName } from '../services/models';
import { DataService } from '../services/DataService';
import { errMsg } from '../services/utils';
import { ErrorBar } from './common';
import { Dashboard } from './Dashboard';
import { AssetList } from './AssetList';
import { AssetForm } from './AssetForm';
import { AssetDetail } from './AssetDetail';
import { AssignReturn } from './AssignReturn';
import { LicenceList } from './LicenceList';
import { LicenceForm } from './LicenceForm';
import { LicenceDetail } from './LicenceDetail';
import { SeatAllocation } from './SeatAllocation';
import { Renewals } from './Renewals';
import { AccessManagement } from './AccessManagement';
import { Settings } from './Settings';

initializeIcons();

export interface IItAssetTrackerProps { appTitle: string; }

const SCREENS: { [k in ViewName]: React.FC<IScreenProps> } = {
  dashboard: Dashboard, assets: AssetList, assetForm: AssetForm, assetDetail: AssetDetail, assignReturn: AssignReturn,
  licences: LicenceList, licenceForm: LicenceForm, licenceDetail: LicenceDetail, seatAllocation: SeatAllocation,
  renewals: Renewals, access: AccessManagement, settings: Settings
};

const SECTION: { [k in ViewName]: string } = {
  dashboard: 'dashboard', assets: 'assets', assetForm: 'assets', assetDetail: 'assets', assignReturn: 'assets',
  licences: 'licences', licenceForm: 'licences', licenceDetail: 'licences', seatAllocation: 'licences',
  renewals: 'renewals', access: 'access', settings: 'settings'
};

const ItAssetTracker: React.FC<IItAssetTrackerProps> = ({ appTitle }) => {
  const [role, setRole] = React.useState<Role | undefined>();
  const [fatal, setFatal] = React.useState<string | undefined>();
  const [warning, setWarning] = React.useState<string | undefined>();
  const [view, setView] = React.useState<IView>({ name: 'dashboard' });

  React.useEffect(() => {
    DataService.onWarning = setWarning;
    DataService.loadIdentityAndRole().then(setRole).catch(e => setFatal(
      `${errMsg(e)} — if the lists do not exist yet, run scripts/Provision-ITAM.ps1 against this site.`));
  }, []);

  if (fatal) { return <ErrorBar error={fatal} />; }
  if (!role) { return <Spinner size={SpinnerSize.large} label="Checking your access…" style={{ padding: 60 }} />; }
  if (role === 'none') { return <ErrorBar error="Access Denied: you are not authorised to use the IT Asset Tracker." />; }

  const go = (v: IView): void => { setView(v); window.scrollTo(0, 0); };
  const Screen = SCREENS[view.name];
  const links = [
    { name: 'Dashboard', key: 'dashboard', icon: 'ViewDashboard' }, { name: 'Assets', key: 'assets', icon: 'Devices2' },
    { name: 'Software licences', key: 'licences', icon: 'Installation' }, { name: 'Renewals & warranties', key: 'renewals', icon: 'Calendar' },
    ...(role === 'admin' ? [{ name: 'User access', key: 'access', icon: 'Permissions' }, { name: 'Settings & audit', key: 'settings', icon: 'Settings' }] : [])
  ].map(l => ({ ...l, url: '', iconProps: { iconName: l.icon }, onClick: () => go({ name: l.key as ViewName }) }));

  return (
    <div style={{ fontFamily: 'inherit' }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ padding: '8px 4px', borderBottom: '1px solid #c8c6c4' }}>
        <Text variant="xLarge" style={{ fontWeight: 600 }}>{appTitle}</Text>
        <Text variant="small">{DataService.me.Title} • {role === 'admin' ? 'Owner / Admin' : role === 'editor' ? 'IT Admin / Editor' : 'Viewer'}</Text>
      </Stack>
      <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginTop: 8 }}>
        <div style={{ width: 210, flex: '0 0 210px' }}>
          <Nav groups={[{ links }]} selectedKey={SECTION[view.name]} ariaLabel="Main navigation" />
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          {warning && <ErrorBar error={warning} type={MessageBarType.severeWarning} onDismiss={() => setWarning(undefined)} />}
          <Screen key={`${view.name}-${view.id || 0}-${JSON.stringify(view.params || {})}`} role={role} view={view} go={go} />
        </div>
      </Stack>
    </div>
  );
};

export default ItAssetTracker;
