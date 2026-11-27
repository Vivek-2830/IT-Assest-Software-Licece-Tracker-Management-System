import { SharePointService } from './SharePointService';
import { IAsset, IAssetFormValues } from '../models/IAsset';
import { sp } from '@pnp/sp/rest';

export const ASSET_LIST_TITLE: string = 'IT_Assets';

const ASSET_SELECT: string[] = [
  'Id', 'Title', 'AssetTag', 'AssetType', 'Manufacturer', 'Model',
  'SerialNumber', 'PurchaseDate', 'PurchaseCost', 'Supplier',
  'InvoiceReference', 'WarrantyEnd', 'OfficeLocation', 'AssetStatus',
  'CurrentHolder/Id', 'CurrentHolder/Title', 'CurrentHolder/EMail',
  'CurrentAssignmentID', 'ArchivedOn', 'Notes'
];

export class AssetService {
  public static async getAssets(includeArchived: boolean = false): Promise<IAsset[]> {
    const list: any = SharePointService.getList(ASSET_LIST_TITLE);
    const rows: any[] = await list.items
      .select.apply(list.items, ASSET_SELECT)
      .expand('CurrentHolder')
      .top(5000)
      .orderBy('Id', false)
      .get();

    const mapped: IAsset[] = rows.map((row: any) => ({
      Id: row.Id,
      Title: row.Title || '',
      AssetTag: row.AssetTag || '',
      AssetType: row.AssetType || '',
      Manufacturer: row.Manufacturer || '',
      Model: row.Model || '',
      SerialNumber: row.SerialNumber || '',
      PurchaseDate: row.PurchaseDate || '',
      PurchaseCost: row.PurchaseCost,
      Supplier: row.Supplier || '',
      InvoiceReference: row.InvoiceReference || '',
      WarrantyEnd: row.WarrantyEnd || '',
      OfficeLocation: row.OfficeLocation || '',
      AssetStatus: row.AssetStatus || 'Available',
      CurrentHolder: row.CurrentHolder ? {
        Id: row.CurrentHolder.Id,
        Title: row.CurrentHolder.Title,
        EMail: row.CurrentHolder.EMail
      } : undefined,
      CurrentAssignmentID: row.CurrentAssignmentID,
      ArchivedOn: row.ArchivedOn || '',
      Notes: row.Notes || ''
    }));

    return includeArchived ? mapped : mapped.filter((asset: IAsset) => !asset.ArchivedOn);
  }

  private static async validateUnique(values: IAssetFormValues, currentId?: number): Promise<void> {
    const list: any = SharePointService.getList(ASSET_LIST_TITLE);
    const assetTag: string = values.AssetTag.trim().replace(/'/g, "''");
    const tagRows: any[] = await list.items
      .filter("AssetTag eq '" + assetTag + "'")
      .select('Id')
      .top(2)
      .get();

    if (tagRows.some((row: any) => row.Id !== currentId)) {
      throw new Error('Asset Tag already exists. Please enter a unique Asset Tag.');
    }

    if (values.SerialNumber && values.SerialNumber.trim()) {
      const serial: string = values.SerialNumber.trim().replace(/'/g, "''");
      const serialRows: any[] = await list.items
        .filter("SerialNumber eq '" + serial + "'")
        .select('Id')
        .top(2)
        .get();

      if (serialRows.some((row: any) => row.Id !== currentId)) {
        throw new Error('Serial Number already exists.');
      }
    }
  }

  // private static async buildPayload(values: IAssetFormValues, isNew: boolean): Promise<any> {
  //   const payload: any = {
  //     Title: values.Title.trim(),
  //     AssetTag: values.AssetTag.trim(),
  //     AssetType: values.AssetType,
  //     Manufacturer: values.Manufacturer || '',
  //     Model: values.Model || '',
  //     SerialNumber: values.SerialNumber || '',
  //     PurchaseDate: values.PurchaseDate ? new Date(values.PurchaseDate).toISOString() : null,
  //     PurchaseCost: values.PurchaseCost ? Number(values.PurchaseCost) : null,
  //     Supplier: values.Supplier || '',
  //     InvoiceReference: values.InvoiceReference || '',
  //     WarrantyEnd: values.WarrantyEnd ? new Date(values.WarrantyEnd).toISOString() : null,
  //     OfficeLocation: values.OfficeLocation || 'Unspecified',
  //     Notes: values.Notes || ''
  //   };

  //   if (values.HolderEmail && values.HolderEmail.trim()) {
  //     payload.CurrentHolderId = await SharePointService.ensureUser(values.HolderEmail);
  //   } else {
  //     payload.CurrentHolderId = null;
  //   }

  //   if (isNew) {
  //     payload.AssetStatus = 'Available';
  //   }
  //   return payload;
  // }

  private static async buildPayload(
    values: IAssetFormValues,
    isNew: boolean
  ): Promise<any> {
    const payload: any = {
      Title: values.Title.trim(),
      AssetTag: values.AssetTag.trim(),
      AssetType: values.AssetType,
      Manufacturer: values.Manufacturer || '',
      Model: values.Model || '',
      SerialNumber: values.SerialNumber || '',
      PurchaseDate: values.PurchaseDate
        ? new Date(values.PurchaseDate).toISOString()
        : null,
      PurchaseCost: values.PurchaseCost
        ? Number(values.PurchaseCost)
        : null,
      Supplier: values.Supplier || '',
      InvoiceReference: values.InvoiceReference || '',
      WarrantyEnd: values.WarrantyEnd
        ? new Date(values.WarrantyEnd).toISOString()
        : null,
      OfficeLocation: values.OfficeLocation || 'Unspecified',
      Notes: values.Notes || ''
    };

    // Save selected SharePoint Person
    const holderEmail = (values.HolderEmail || '').trim();

    if (holderEmail) {
      payload.CurrentHolderId =
        await SharePointService.ensureUser(holderEmail);
    } else {
      payload.CurrentHolderId = null;
    }

    if (isNew) {
      payload.AssetStatus = 'Available';
    }

    return payload;
  }

  public static async ensureUser(email: string): Promise<number> {
    const result = await sp.web.ensureUser(email);
    return result.data.Id;
  }

  public static async createAsset(values: IAssetFormValues): Promise<void> {
    await this.validateUnique(values);
    const payload: any = await this.buildPayload(values, true);
    await SharePointService.getList(ASSET_LIST_TITLE).items.add(payload);
  }

  public static async updateAsset(id: number, values: IAssetFormValues): Promise<void> {
    await this.validateUnique(values, id);
    const payload: any = await this.buildPayload(values, false);
    await SharePointService.getList(ASSET_LIST_TITLE).items.getById(id).update(payload);
  }

  public static async archiveAsset(id: number): Promise<void> {
    await SharePointService.getList(ASSET_LIST_TITLE).items.getById(id).update({
      ArchivedOn: new Date().toISOString()
    });
  }

  public static async getDashboardMetrics(): Promise<{
    totalAssets: number;
    availableAssets: number;
    assignedAssets: number;
    repairAssets: number;
    retiredAssets: number;
    warrantiesDue: number;
  }> {
    const assets: IAsset[] = await this.getAssets(false);
    const today: Date = new Date();
    const end: Date = new Date();
    end.setDate(end.getDate() + 30);

    const warrantiesDue: number = assets.filter((asset: IAsset) => {
      if (!asset.WarrantyEnd) { return false; }
      const warrantyDate: Date = new Date(asset.WarrantyEnd);
      return warrantyDate >= today && warrantyDate <= end;
    }).length;

    return {
      totalAssets: assets.length,
      availableAssets: assets.filter((a: IAsset) => a.AssetStatus === 'Available').length,
      assignedAssets: assets.filter((a: IAsset) => a.AssetStatus === 'Assigned').length,
      repairAssets: assets.filter((a: IAsset) => a.AssetStatus === 'Repair').length,
      retiredAssets: assets.filter((a: IAsset) => a.AssetStatus === 'Retired').length,
      warrantiesDue: warrantiesDue
    };
  }
}
