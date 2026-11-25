export const LISTS = {
  Assets: 'IT_Assets',
  AssetAssignments: 'Asset_Assignments',
  Licences: 'Software_Licences',
  LicenceAssignments: 'Licence_Assignments',
  AccessRequests: 'Access_Requests',
  ActivityLog: 'Activity_Log',
  Settings: 'App_Settings',
  ReminderLog: 'Reminder_Log',
  Documents: 'ITAM_Documents'
};

export const ASSET_TYPES = ['Laptop', 'Desktop', 'Monitor', 'Phone', 'Tablet', 'Network', 'Accessory', 'Other'];
export const LOCATIONS = ['UK', 'India', 'Other', 'Unspecified'];
export const ASSET_STATUSES = ['Available', 'Assigned', 'Repair', 'Retired'];
export const CONDITIONS = ['Good', 'Fair', 'Damaged', 'Unknown'];
export const LICENCE_MODELS = ['Named-user seat', 'Device seat', 'Concurrent seat', 'Site/Unlimited', 'Non-seat', 'Perpetual'];
export const SEAT_MODELS = ['Named-user seat', 'Device seat', 'Concurrent seat'];
export const BILLING = ['Monthly', 'Quarterly', 'Yearly', 'One-time', 'Other'];
export const CURRENCIES = ['GBP', 'USD', 'INR', 'EUR'];
export const LICENCE_STATUSES = ['Active', 'Cancelled', 'Expired', 'Archived'];
export const ACCESS_ROLES = ['Admin', 'Editor', 'Viewer'];
export const ACCESS_STATUSES = ['Pending', 'Approved', 'Active', 'Rejected', 'Revoked'];
export const DOC_TYPES = ['Invoice', 'Warranty', 'Contract', 'Other'];
export const ENTITY_TYPES = ['Asset', 'Assignment', 'Licence', 'Licence Allocation', 'Access', 'Settings'];

/** Spec section 5 / 13 say GBP for asset cost and as the default currency (section 1.3 says INR). Change here if needed. */
export const DEFAULT_CURRENCY = 'GBP';
export const ASSET_TAG_PREFIX = 'IT-';
export const LICENCE_CODE_PREFIX = 'LIC-';
export const DEFAULT_REMINDER_DAYS = '30,15,7';
