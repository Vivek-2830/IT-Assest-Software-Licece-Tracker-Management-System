import { SharePointService } from './SharePointService';
import { AssetService } from './AssetService';
import { IAssetDashboardMetrics } from '../models/IAsset';

export class DashboardService {
  public static async getMetrics(): Promise<IAssetDashboardMetrics> {
    const assetMetrics: any = await AssetService.getDashboardMetrics();
    let activeLicences: number | null = null;

    try {
      const list: any = SharePointService.getList('Software_Licences');
      const rows: any[] = await list.items
        .select('Id', 'LicenceStatus')
        .filter("LicenceStatus eq 'Active'")
        .top(5000)
        .get();
      activeLicences = rows.length;
    } catch (error) {
      // Licence list may not yet exist; dashboard's asset metrics remain available.
      activeLicences = null;
    }

    return {
      totalAssets: assetMetrics.totalAssets,
      availableAssets: assetMetrics.availableAssets,
      assignedAssets: assetMetrics.assignedAssets,
      repairAssets: assetMetrics.repairAssets,
      retiredAssets: assetMetrics.retiredAssets,
      warrantiesDue: assetMetrics.warrantiesDue,
      activeLicences: activeLicences
    };
  }
}
