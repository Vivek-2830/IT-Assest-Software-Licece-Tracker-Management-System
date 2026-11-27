import * as React from 'react';
import { MessageBar, MessageBarType, Spinner, SpinnerSize, Icon } from '@fluentui/react';
import { DashboardService } from '../../services/DashboardService';
import { IAssetDashboardMetrics } from '../../models/IAsset';
import styles from './Dashboard.module.scss';

const initialMetrics: IAssetDashboardMetrics = {
  totalAssets: 0, availableAssets: 0, assignedAssets: 0,
  repairAssets: 0, retiredAssets: 0, activeLicences: null, warrantiesDue: 0
};

export const Dashboard: React.FC = () => {
  const [metrics, setMetrics] = React.useState<IAssetDashboardMetrics>(initialMetrics);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');

  const load = React.useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      setMetrics(await DashboardService.getMetrics());
    } catch (e) {
      setError(e && e.message ? e.message : 'Unable to load dashboard. Check SharePoint list names and permissions.');
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => { load(); }, [load]);

  const cards: any[] = [
    { title: 'Total Assets', value: metrics.totalAssets, icon: 'Devices', color: '#2563eb' },
    { title: 'Available', value: metrics.availableAssets, icon: 'CheckMark', color: '#16834a' },
    { title: 'Assigned', value: metrics.assignedAssets, icon: 'Contact', color: '#7c3aed' },
    { title: 'In Repair', value: metrics.repairAssets, icon: 'Repair', color: '#d97706' },
    { title: 'Active Licences', value: metrics.activeLicences === null ? '—' : metrics.activeLicences, icon: 'Key', color: '#0891b2' },
    { title: 'Warranty Due (30 days)', value: metrics.warrantiesDue, icon: 'Calendar', color: '#dc5b42' }
  ];

  return (
    <div className={styles.dashboard}>
      <div className={styles.sectionHeading}>
        <div><h2>Overview</h2><p>Current asset and licence position</p></div>
        <button className={styles.refreshButton} onClick={load} type="button">Refresh</button>
      </div>

      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

      {loading ? (
        <div className={styles.loading}><Spinner size={SpinnerSize.large} label="Loading dashboard..." /></div>
      ) : (
        <div className={styles.metricGrid}>
          {cards.map((card: any) => (
            <div className={styles.metricCard} key={card.title}>
              <div className={styles.metricTop}>
                <span className={styles.metricIcon} style={{ color: card.color }}><Icon iconName={card.icon} /></span>
                <span className={styles.metricLabel}>{card.title}</span>
              </div>
              <div className={styles.metricValue}>{card.value}</div>
            </div>
          ))}
        </div>
      )}

      <div className={styles.bottomGrid}>
        <div className={styles.panel}>
          <h3>Asset status</h3>
          <div className={styles.statusRow}><span>Available</span><strong>{metrics.availableAssets}</strong></div>
          <div className={styles.statusRow}><span>Assigned</span><strong>{metrics.assignedAssets}</strong></div>
          <div className={styles.statusRow}><span>Repair</span><strong>{metrics.repairAssets}</strong></div>
          <div className={styles.statusRow}><span>Retired</span><strong>{metrics.retiredAssets}</strong></div>
        </div>
        <div className={styles.panel}>
          <h3>Attention required</h3>
          <div className={styles.attentionNumber}>{metrics.warrantiesDue}</div>
          <p>Assets with warranty expiry in the next 30 days.</p>
        </div>
      </div>
    </div>
  );
};
