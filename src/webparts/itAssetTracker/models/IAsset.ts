export type AssetStatus = 'Available' | 'Assigned' | 'Repair' | 'Retired';

export interface IAsset {
  Id: number;
  Title: string;
  AssetTag: string;
  AssetType: string;
  Manufacturer?: string;
  Model?: string;
  SerialNumber?: string;
  PurchaseDate?: string;
  PurchaseCost?: number;
  Supplier?: string;
  InvoiceReference?: string;
  WarrantyEnd?: string;
  OfficeLocation?: string;
  AssetStatus: AssetStatus;
  CurrentHolder?: {
    Id?: number;
    Title?: string;
    EMail?: string;
  };
  CurrentAssignmentID?: number;
  ArchivedOn?: string;
  Notes?: string;
}

export interface IAssetFormValues {
  Title: string;
  AssetTag: string;
  AssetType: string;
  Manufacturer: string;
  Model: string;
  SerialNumber: string;
  PurchaseDate: string;
  PurchaseCost: string;
  Supplier: string;
  InvoiceReference: string;
  WarrantyEnd: string;
  OfficeLocation: string;
  HolderEmail: string;
  Notes: string;
}

export interface IAssetDashboardMetrics {
  totalAssets: number;
  availableAssets: number;
  assignedAssets: number;
  repairAssets: number;
  retiredAssets: number;
  activeLicences: number | null;
  warrantiesDue: number;
}
