import * as React from 'react';
import { Pivot, PivotItem, MessageBar, MessageBarType,  NormalPeoplePicker, IPersonaProps } from '@fluentui/react';
import { Dashboard } from './dashboard/Dashboard';
import { Assets } from './assets/Assets';
import styles from './ItAssetTracker.module.scss';
import { IITAssetTrackerProps } from './IItAssetTrackerProps';

const ITAssetTracker: React.FC<IITAssetTrackerProps> = (props: IITAssetTrackerProps) => {
  const [error, setError] = React.useState<string>('');
  const [selectedHolder, setSelectedHolder] = React.useState<IPersonaProps[]>([]);

  return (
    <div className={styles.itAssetTracker}>
      <div className={styles.header}>
        <div>
          <div className={styles.eyebrow}>IT OPERATIONS</div>
          <h1>IT Asset Tracker</h1>
          <p>Manage company assets and software licences in one place.</p>
        </div>
      </div>

      {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>{error}</MessageBar>}

      <Pivot aria-label="IT Asset Tracker navigation" className={styles.pivot}>
        <PivotItem headerText="Dashboard" itemKey="dashboard">
          <Dashboard />
        </PivotItem>
        <PivotItem headerText="IT Assets" itemKey="assets">
          <Assets />
        </PivotItem>
        <PivotItem headerText="Software Licences" itemKey="licences">
          <div className={styles.placeholder}>Software Licences module will be implemented in the next phase.</div>
        </PivotItem>
        <PivotItem headerText="Access Requests" itemKey="access">
          <div className={styles.placeholder}>Access Requests module will be implemented in the next phase.</div>
        </PivotItem>
        <PivotItem headerText="Activity Log" itemKey="activity">
          <div className={styles.placeholder}>Activity Log module will be implemented in the next phase.</div>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default ITAssetTracker;
