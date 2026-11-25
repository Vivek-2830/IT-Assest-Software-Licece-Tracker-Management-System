export type Role = 'admin' | 'editor' | 'viewer' | 'none';

export type ViewName =
  | 'dashboard' | 'assets' | 'assetForm' | 'assetDetail' | 'assignReturn'
  | 'licences' | 'licenceForm' | 'licenceDetail' | 'seatAllocation'
  | 'renewals' | 'access' | 'settings';

export interface IView { name: ViewName; id?: number; params?: any; }

export interface IScreenProps {
  role: Role;
  view: IView;
  go: (v: IView) => void;
}

export interface IPerson { Id: number; Title: string; EMail?: string; }

export interface IAsset {
  Id: number; Title: string; AssetTag: string; AssetType: string;
  Manufacturer?: string; Model?: string; SerialNumber?: string;
  PurchaseDate?: Date; PurchaseCost?: number; Supplier?: string; InvoiceReference?: string;
  WarrantyEnd?: Date; OfficeLocation: string; AssetStatus: string;
  CurrentHolder?: IPerson; CurrentAssignmentID?: number; Notes?: string; ArchivedOn?: Date;
  Modified: string;
}

export interface IAssetInput {
  Title: string; AssetTag: string; AssetType: string; Manufacturer?: string; Model?: string;
  SerialNumber?: string; PurchaseDate?: Date; PurchaseCost?: number; Supplier?: string;
  InvoiceReference?: string; WarrantyEnd?: Date; OfficeLocation: string; Notes?: string;
}

export interface IAssetAssignment {
  Id: number; AssetID: number; AssetTagSnapshot: string; AssignedTo?: IPerson;
  AssignedDate?: Date; ReturnedDate?: Date; ConditionOut?: string; ConditionIn?: string;
  AssignedBy?: IPerson; ReceivedBy?: IPerson; AssignmentStatus: string; Notes?: string;
}

export interface ILicence {
  Id: number; Title: string; LicenceCode: string; Vendor: string; LicenceModel: string;
  PurchasedSeats?: number; StartDate?: Date; RenewalDate?: Date; Renewable: boolean; AutoRenew: boolean;
  BillingFrequency: string; Amount?: number; CurrencyCode: string; LicenceStatus: string;
  VendorContact?: string; ContractReference?: string; Notes?: string; LastRenewedOn?: Date;
  Modified: string;
}

export interface ILicenceInput {
  Title: string; LicenceCode: string; Vendor: string; LicenceModel: string; PurchasedSeats?: number;
  StartDate?: Date; RenewalDate?: Date; Renewable: boolean; AutoRenew: boolean;
  BillingFrequency: string; Amount?: number; CurrencyCode: string; LicenceStatus: string;
  VendorContact?: string; ContractReference?: string; Notes?: string;
}

export interface ILicenceAllocation {
  Id: number; LicenceID: number; LicenceCodeSnapshot: string; TargetPerson?: IPerson;
  TargetAssetID?: number; TargetAssetTag?: string; AssignedDate?: Date; ReclaimedDate?: Date;
  AllocationStatus: string; Notes?: string;
}

export interface IAccessRequest {
  Id: number; Title: string; UserEmail: string; RequestedRole: string; AccessStatus: string;
  RequestedBy?: IPerson; ApprovedBy?: IPerson; ProvisionedBy?: IPerson;
  RequestedOn?: string; ApprovedOn?: string; ProvisionedOn?: string; Notes?: string;
}

export interface IActivity {
  Id: number; Title: string; EntityType: string; EntityID?: number; ActionName: string;
  PerformedBy?: IPerson; ActionTime: string; BeforeSummary?: string; AfterSummary?: string;
  Outcome: string; CorrelationID?: string;
}

export interface IDocument { Id: number; FileLeafRef: string; FileRef: string; DocumentType?: string; }
