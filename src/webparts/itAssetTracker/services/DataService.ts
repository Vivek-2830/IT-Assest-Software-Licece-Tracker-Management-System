import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/security';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { PermissionKind } from '@pnp/sp/security';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IAccessRequest, IActivity, IAsset, IAssetAssignment, IAssetInput, IDocument, ILicence,
  ILicenceAllocation, ILicenceInput, IPerson, Role
} from './models';
import {
  ASSET_TAG_PREFIX, LICENCE_CODE_PREFIX, LISTS, SEAT_MODELS
} from './constants';
import { esc, fmt, guid, nextCode, parseDate, sameOrAfter, toDateOnly, today } from './utils';

const P = (r: any): IPerson | undefined => (r && r.Id ? { Id: r.Id, Title: r.Title, EMail: r.EMail } : undefined);
const num = (v: any): number | undefined => (v === null || v === undefined || v === '' ? undefined : Number(v));

const ASSET_SEL = ['Id', 'Title', 'AssetTag', 'AssetType', 'Manufacturer', 'Model', 'SerialNumber', 'PurchaseDate',
  'PurchaseCost', 'Supplier', 'InvoiceReference', 'WarrantyEnd', 'OfficeLocation', 'AssetStatus',
  'CurrentHolder/Id', 'CurrentHolder/Title', 'CurrentHolder/EMail', 'CurrentAssignmentID', 'Notes', 'ArchivedOn', 'Modified'];
const AA_SEL = ['Id', 'AssetID', 'AssetTagSnapshot', 'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail', 'AssignedDate',
  'ReturnedDate', 'ConditionOut', 'ConditionIn', 'AssignedBy/Id', 'AssignedBy/Title', 'ReceivedBy/Id', 'ReceivedBy/Title',
  'AssignmentStatus', 'Notes'];
const LIC_SEL = ['Id', 'Title', 'LicenceCode', 'Vendor', 'LicenceModel', 'PurchasedSeats', 'StartDate', 'RenewalDate', 'Renewable',
  'AutoRenew', 'BillingFrequency', 'Amount', 'CurrencyCode', 'LicenceStatus', 'VendorContact', 'ContractReference', 'Notes',
  'LastRenewedOn', 'Modified'];
const LA_SEL = ['Id', 'LicenceID', 'LicenceCodeSnapshot', 'TargetPerson/Id', 'TargetPerson/Title', 'TargetPerson/EMail',
  'TargetAssetID', 'AssignedDate', 'ReclaimedDate', 'AllocationStatus', 'Notes'];
const AR_SEL = ['Id', 'Title', 'UserEmail', 'RequestedRole', 'AccessStatus', 'RequestedBy/Id', 'RequestedBy/Title',
  'ApprovedBy/Id', 'ApprovedBy/Title', 'ProvisionedBy/Id', 'ProvisionedBy/Title', 'RequestedOn', 'ApprovedOn', 'ProvisionedOn', 'Notes'];
const LOG_SEL = ['Id', 'Title', 'EntityType', 'EntityID', 'ActionName', 'PerformedBy/Id', 'PerformedBy/Title', 'ActionTime',
  'BeforeSummary', 'AfterSummary', 'Outcome', 'CorrelationID'];

const mapAsset = (r: any): IAsset => ({
  Id: r.Id, Title: r.Title, AssetTag: r.AssetTag, AssetType: r.AssetType, Manufacturer: r.Manufacturer || undefined,
  Model: r.Model || undefined, SerialNumber: r.SerialNumber || undefined, PurchaseDate: parseDate(r.PurchaseDate),
  PurchaseCost: num(r.PurchaseCost), Supplier: r.Supplier || undefined, InvoiceReference: r.InvoiceReference || undefined,
  WarrantyEnd: parseDate(r.WarrantyEnd), OfficeLocation: r.OfficeLocation, AssetStatus: r.AssetStatus,
  CurrentHolder: P(r.CurrentHolder), CurrentAssignmentID: num(r.CurrentAssignmentID), Notes: r.Notes || undefined,
  ArchivedOn: parseDate(r.ArchivedOn), Modified: r.Modified
});
const mapAA = (r: any): IAssetAssignment => ({
  Id: r.Id, AssetID: r.AssetID, AssetTagSnapshot: r.AssetTagSnapshot, AssignedTo: P(r.AssignedTo),
  AssignedDate: parseDate(r.AssignedDate), ReturnedDate: parseDate(r.ReturnedDate), ConditionOut: r.ConditionOut,
  ConditionIn: r.ConditionIn, AssignedBy: P(r.AssignedBy), ReceivedBy: P(r.ReceivedBy),
  AssignmentStatus: r.AssignmentStatus, Notes: r.Notes
});
const mapLic = (r: any): ILicence => ({
  Id: r.Id, Title: r.Title, LicenceCode: r.LicenceCode, Vendor: r.Vendor, LicenceModel: r.LicenceModel,
  PurchasedSeats: num(r.PurchasedSeats), StartDate: parseDate(r.StartDate), RenewalDate: parseDate(r.RenewalDate),
  Renewable: !!r.Renewable, AutoRenew: !!r.AutoRenew, BillingFrequency: r.BillingFrequency, Amount: num(r.Amount),
  CurrencyCode: r.CurrencyCode, LicenceStatus: r.LicenceStatus, VendorContact: r.VendorContact || undefined,
  ContractReference: r.ContractReference || undefined, Notes: r.Notes || undefined,
  LastRenewedOn: parseDate(r.LastRenewedOn), Modified: r.Modified
});
const mapLA = (r: any): ILicenceAllocation => ({
  Id: r.Id, LicenceID: r.LicenceID, LicenceCodeSnapshot: r.LicenceCodeSnapshot, TargetPerson: P(r.TargetPerson),
  TargetAssetID: num(r.TargetAssetID), AssignedDate: parseDate(r.AssignedDate), ReclaimedDate: parseDate(r.ReclaimedDate),
  AllocationStatus: r.AllocationStatus, Notes: r.Notes
});

export const isSeatManaged = (l: ILicence): boolean =>
  SEAT_MODELS.indexOf(l.LicenceModel) >= 0 || (l.LicenceModel === 'Perpetual' && !!l.PurchasedSeats && l.PurchasedSeats > 0);

class DataServiceImpl {
  private ctx!: WebPartContext;
  public me: IPerson = { Id: 0, Title: '' };
  /** Set by the shell: surfaces non-fatal problems (e.g. failed audit writes) to the user. */
  public onWarning?: (msg: string) => void;

  public init(ctx: WebPartContext): void {
    this.ctx = ctx;
    sp.setup({ spfxContext: ctx as any });
  }

  private items(list: string): any { return sp.web.lists.getByTitle(list).items; }
  private warn(m: string): void { if (this.onWarning) { this.onWarning(m); } }

  // ---------------------------------------------------------------- identity / roles
  /**
   * Role is derived from REAL SharePoint permissions (which are granted through the Entra groups
   * ITAM-Admins / ITAM-Editors / ITAM-Viewers). UI hiding is only a convenience; SharePoint enforces access.
   *   Admin  = Manage Web (Full Control on the site)   Editor = Edit items   Viewer = View items
   */
  public async loadIdentityAndRole(): Promise<Role> {
    const u: any = await sp.web.currentUser.select('Id', 'Title', 'Email').get();
    this.me = { Id: u.Id, Title: u.Title, EMail: u.Email };
    const assets = sp.web.lists.getByTitle(LISTS.Assets);
    const [admin, edit, view] = await Promise.all([
      sp.web.currentUserHasPermissions(PermissionKind.ManageWeb),
      assets.currentUserHasPermissions(PermissionKind.EditListItems),
      assets.currentUserHasPermissions(PermissionKind.ViewListItems)
    ]);
    return admin ? 'admin' : edit ? 'editor' : view ? 'viewer' : 'none';
  }

  public async searchPeople(q: string): Promise<Array<{ Key: string; DisplayText: string; Email?: string }>> {
    const url = `${this.ctx.pageContext.web.absoluteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const body = JSON.stringify({
      queryParams: {
        __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
        QueryString: q, MaximumEntitySuggestions: 10, AllowEmailAddresses: true, AllowOnlyEmailAddresses: false,
        PrincipalType: 1, PrincipalSource: 15, SharePointGroupID: 0
      }
    });
    const res = await this.ctx.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'odata-version': '3.0' },
      body
    });
    const json = await res.json();
    const arr: any[] = JSON.parse(json.d.ClientPeoplePickerSearchUser);
    return arr.map(x => ({ Key: x.Key, DisplayText: x.DisplayText, Email: x.EntityData && x.EntityData.Email }));
  }

  /** Resolve a directory person to a site user (employees are custodians only; this implies no app access). */
  public async ensureUser(login: string): Promise<IPerson> {
    const r: any = await sp.web.ensureUser(login);
    return { Id: r.data.Id, Title: r.data.Title, EMail: r.data.Email };
  }

  // ---------------------------------------------------------------- activity log
  public async log(entityType: string, entityId: number | undefined, action: string, before?: string, after?: string,
    outcome: 'Success' | 'Failure' = 'Success', title?: string): Promise<void> {
    try {
      await this.items(LISTS.ActivityLog).add({
        Title: title || `${action} ${entityType}`, EntityType: entityType, EntityID: entityId, ActionName: action,
        PerformedById: this.me.Id, ActionTime: new Date().toISOString(), BeforeSummary: before || null,
        AfterSummary: after || null, Outcome: outcome, CorrelationID: guid()
      });
    } catch (e) {
      this.warn(`Audit log write failed for "${action} ${entityType} ${entityId || ''}". Review in Settings > Audit. SharePoint version history remains the fallback.`);
    }
  }

  public async getLog(top: number, filter?: string): Promise<IActivity[]> {
    let q = this.items(LISTS.ActivityLog).select(...LOG_SEL).expand('PerformedBy').orderBy('Id', false).top(top);
    if (filter) { q = q.filter(filter); }
    const r: any[] = await q.get();
    return r.map(x => ({ ...x, PerformedBy: P(x.PerformedBy) } as IActivity));
  }

  private diff(a: any, b: any, keys: string[]): string {
    const out: string[] = [];
    keys.forEach(k => {
      const x = a[k] instanceof Date ? fmt(a[k]) : a[k];
      const y = b[k] instanceof Date ? fmt(b[k]) : b[k];
      if ((x || '') !== (y || '')) { out.push(`${k}: ${x === undefined || x === '' ? '∅' : x} → ${y === undefined || y === '' ? '∅' : y}`); }
    });
    return out.join('; ') || 'no field changes';
  }

  // ---------------------------------------------------------------- assets
  public async getAssets(): Promise<IAsset[]> {
    const r: any[] = await this.items(LISTS.Assets).select(...ASSET_SEL).expand('CurrentHolder').getAll();
    return r.map(mapAsset);
  }

  public async getAsset(id: number): Promise<IAsset> {
    const r = await this.items(LISTS.Assets).getById(id).select(...ASSET_SEL).expand('CurrentHolder').get();
    return mapAsset(r);
  }

  private async checkAssetDuplicates(tag: string, serial: string | undefined, excludeId?: number): Promise<void> {
    const ex = excludeId ? ` and Id ne ${excludeId}` : '';
    const t: any[] = await this.items(LISTS.Assets).select('Id', 'AssetTag', 'Title').filter(`AssetTag eq '${esc(tag)}'${ex}`).top(1).get();
    if (t.length) { throw new Error(`Duplicate asset tag (V02): ${t[0].AssetTag} "${t[0].Title}" already exists (ID ${t[0].Id}).`); }
    if (serial) {
      const s: any[] = await this.items(LISTS.Assets).select('Id', 'AssetTag', 'Title').filter(`SerialNumber eq '${esc(serial)}'${ex}`).top(1).get();
      if (s.length) { throw new Error(`Duplicate serial number (V03): already recorded on ${s[0].AssetTag} "${s[0].Title}" (ID ${s[0].Id}). Correct the existing row if it is wrong.`); }
    }
  }

  private assetPayload(f: IAssetInput): any {
    return {
      Title: f.Title, AssetTag: f.AssetTag, AssetType: f.AssetType, Manufacturer: f.Manufacturer || null, Model: f.Model || null,
      SerialNumber: f.SerialNumber || null, PurchaseDate: toDateOnly(f.PurchaseDate),
      PurchaseCost: f.PurchaseCost === undefined ? null : f.PurchaseCost, Supplier: f.Supplier || null,
      InvoiceReference: f.InvoiceReference || null, WarrantyEnd: toDateOnly(f.WarrantyEnd), OfficeLocation: f.OfficeLocation,
      Notes: f.Notes || null
    };
  }

  /** WF01 */
  public async createAsset(f: IAssetInput): Promise<number> {
    try {
      if (!f.AssetTag) {
        const tags: any[] = await this.items(LISTS.Assets).select('AssetTag').getAll();
        f.AssetTag = nextCode(ASSET_TAG_PREFIX, tags.map(t => t.AssetTag));
      }
      await this.checkAssetDuplicates(f.AssetTag, f.SerialNumber);
      const res: any = await this.items(LISTS.Assets).add({ ...this.assetPayload(f), AssetStatus: 'Available' });
      await this.log('Asset', res.data.Id, 'Create', undefined, `${f.AssetTag} ${f.Title}`);
      return res.data.Id;
    } catch (e) {
      await this.log('Asset', undefined, 'Create', undefined, `${f.AssetTag} ${f.Title}`, 'Failure');
      throw e;
    }
  }

  public async updateAsset(id: number, f: IAssetInput, loadedModified: string): Promise<void> {
    const cur = await this.getAsset(id);
    if (cur.AssetStatus === 'Retired') { throw new Error('Retired records are read-only. An Owner must restore the asset first.'); }
    if (cur.Modified !== loadedModified) { throw new Error('This asset was changed by someone else while you were editing (V14). Reload the record and re-apply your changes.'); }
    await this.checkAssetDuplicates(f.AssetTag, f.SerialNumber, id);
    await this.items(LISTS.Assets).getById(id).update(this.assetPayload(f));
    await this.log('Asset', id, 'Update', undefined, this.diff(cur, f, ['Title', 'AssetTag', 'AssetType', 'Manufacturer', 'Model', 'SerialNumber', 'PurchaseDate', 'PurchaseCost', 'Supplier', 'InvoiceReference', 'WarrantyEnd', 'OfficeLocation']));
  }

  public async getAssetAssignments(assetId: number): Promise<IAssetAssignment[]> {
    const r: any[] = await this.items(LISTS.AssetAssignments).select(...AA_SEL).expand('AssignedTo', 'AssignedBy', 'ReceivedBy')
      .filter(`AssetID eq ${assetId}`).orderBy('Id', false).getAll();
    return r.map(mapAA);
  }

  private async openAssignments(assetId: number): Promise<IAssetAssignment[]> {
    return (await this.getAssetAssignments(assetId)).filter(a => a.AssignmentStatus === 'Open').sort((a, b) => a.Id - b.Id);
  }

  /** WF02 – authoritative assign. Re-reads state, writes, then re-verifies there is a single open assignment. */
  public async assignAsset(assetId: number, person: IPerson, date: Date, condition: string, comments?: string): Promise<void> {
    const asset = await this.getAsset(assetId);
    if (asset.AssetStatus !== 'Available') { throw new Error(`Cannot assign (V06): asset is currently ${asset.AssetStatus}. Refresh the record.`); }
    if ((await this.openAssignments(assetId)).length) { throw new Error('Cannot assign (V06): the asset already has an open assignment. Run the reconciliation view.'); }
    let created: any;
    try {
      created = await this.items(LISTS.AssetAssignments).add({
        Title: `${asset.AssetTag} ${toDateOnly(date)}`, AssetID: assetId, AssetTagSnapshot: asset.AssetTag, AssignedToId: person.Id,
        AssignedDate: toDateOnly(date), ConditionOut: condition, AssignedById: this.me.Id, AssignmentStatus: 'Open', Notes: comments || null
      });
    } catch (e) {
      await this.log('Assignment', assetId, 'Assign', undefined, `${asset.AssetTag} → ${person.Title}`, 'Failure');
      throw e;
    }
    const newId: number = created.data.Id;
    const open = await this.openAssignments(assetId);
    if (open.length > 1 && open[0].Id !== newId) {
      await this.items(LISTS.AssetAssignments).getById(newId).update({
        AssignmentStatus: 'Closed', ReturnedDate: toDateOnly(date), Notes: 'Auto-cancelled: concurrent assignment conflict'
      });
      await this.log('Assignment', assetId, 'Assign', undefined, 'Rejected: concurrent assignment', 'Failure');
      throw new Error('Another assignment for this asset was saved at the same time (V14). Your assignment was cancelled. Refresh and retry.');
    }
    try {
      await this.items(LISTS.Assets).getById(assetId).update({
        CurrentHolderId: person.Id, CurrentAssignmentID: newId, AssetStatus: 'Assigned'
      });
    } catch (e) {
      await this.log('Assignment', assetId, 'Assign', undefined, `PARTIAL WRITE: assignment ${newId} created but asset not updated`, 'Failure');
      throw new Error(`Partial write: assignment ${newId} was created but the asset record was not updated. Do NOT retry; ask an administrator to reconcile (Settings > Reconciliation).`);
    }
    await this.log('Assignment', assetId, 'Assign', 'Available', `${asset.AssetTag} → ${person.Title} (${condition})`);
  }

  /** WF03 – return. nextStatus: Available or Repair. */
  public async returnAsset(assetId: number, date: Date, condition: string, nextStatus: string, comments?: string): Promise<void> {
    const asset = await this.getAsset(assetId);
    if (asset.AssetStatus !== 'Assigned') { throw new Error(`Cannot return (V07): asset is ${asset.AssetStatus}, not Assigned.`); }
    const open = await this.openAssignments(assetId);
    if (!open.length) { throw new Error('Cannot return (V07): no open assignment exists. Run the reconciliation view.'); }
    const a = open[open.length - 1];
    if (a.AssignedDate && !sameOrAfter(date, a.AssignedDate)) { throw new Error(`Return date cannot precede the assignment date ${fmt(a.AssignedDate)} (V08).`); }
    await this.items(LISTS.AssetAssignments).getById(a.Id).update({
      ReturnedDate: toDateOnly(date), ConditionIn: condition, ReceivedById: this.me.Id, AssignmentStatus: 'Closed',
      Notes: [a.Notes, comments].filter(x => !!x).join(' | ') || null
    });
    try {
      await this.items(LISTS.Assets).getById(assetId).update({ CurrentHolderId: null, CurrentAssignmentID: null, AssetStatus: nextStatus });
    } catch (e) {
      await this.log('Assignment', assetId, 'Return', undefined, `PARTIAL WRITE: assignment ${a.Id} closed but asset not updated`, 'Failure');
      throw new Error(`Partial write: assignment ${a.Id} was closed but the asset record was not updated. Ask an administrator to reconcile.`);
    }
    await this.log('Assignment', assetId, 'Return', 'Assigned', `${asset.AssetTag} returned (${condition}) → ${nextStatus}`);
  }

  /** Repair / Available / Retire / Restore transitions. Only legal moves are allowed. */
  public async changeAssetStatus(assetId: number, target: string, isAdmin: boolean): Promise<void> {
    const a = await this.getAsset(assetId);
    const allowed: { [k: string]: string[] } = { Available: ['Repair', 'Retired'], Repair: ['Available', 'Retired'], Retired: ['Available'], Assigned: [] };
    if (allowed[a.AssetStatus].indexOf(target) < 0) {
      throw new Error(a.AssetStatus === 'Assigned' ? 'Return the asset first; an assigned asset cannot change status (V06).' : `Cannot move from ${a.AssetStatus} to ${target}.`);
    }
    if (a.AssetStatus === 'Retired' && !isAdmin) { throw new Error('Only an Owner can restore a retired asset.'); }
    if ((await this.openAssignments(assetId)).length) { throw new Error('The asset still has an open assignment; return it first.'); }
    await this.items(LISTS.Assets).getById(assetId).update({
      AssetStatus: target, ArchivedOn: target === 'Retired' ? toDateOnly(today()) : null
    });
    await this.log('Asset', assetId, target === 'Retired' ? 'Retire' : a.AssetStatus === 'Retired' ? 'Restore' : 'Status', a.AssetStatus, target);
  }

  // ---------------------------------------------------------------- licences
  public async getLicences(): Promise<ILicence[]> {
    return (await this.items(LISTS.Licences).select(...LIC_SEL).getAll()).map(mapLic);
  }
  public async getLicence(id: number): Promise<ILicence> {
    return mapLic(await this.items(LISTS.Licences).getById(id).select(...LIC_SEL).get());
  }

  private licPayload(f: ILicenceInput): any {
    return {
      Title: f.Title, LicenceCode: f.LicenceCode, Vendor: f.Vendor, LicenceModel: f.LicenceModel,
      PurchasedSeats: f.PurchasedSeats === undefined ? null : f.PurchasedSeats, StartDate: toDateOnly(f.StartDate),
      RenewalDate: toDateOnly(f.RenewalDate), Renewable: f.Renewable, AutoRenew: f.Renewable ? f.AutoRenew : false,
      BillingFrequency: f.BillingFrequency, Amount: f.Amount === undefined ? null : f.Amount, CurrencyCode: f.CurrencyCode,
      LicenceStatus: f.LicenceStatus, VendorContact: f.VendorContact || null, ContractReference: f.ContractReference || null,
      Notes: f.Notes || null
    };
  }

  private async checkLicenceCode(code: string, excludeId?: number): Promise<void> {
    const ex = excludeId ? ` and Id ne ${excludeId}` : '';
    const t: any[] = await this.items(LISTS.Licences).select('Id', 'LicenceCode', 'Title').filter(`LicenceCode eq '${esc(code)}'${ex}`).top(1).get();
    if (t.length) { throw new Error(`Duplicate licence code (V02): ${t[0].LicenceCode} "${t[0].Title}" already exists (ID ${t[0].Id}).`); }
  }

  public async createLicence(f: ILicenceInput): Promise<number> {
    if (!f.LicenceCode) {
      const codes: any[] = await this.items(LISTS.Licences).select('LicenceCode').getAll();
      f.LicenceCode = nextCode(LICENCE_CODE_PREFIX, codes.map(c => c.LicenceCode));
    }
    await this.checkLicenceCode(f.LicenceCode);
    const res: any = await this.items(LISTS.Licences).add(this.licPayload(f));
    await this.log('Licence', res.data.Id, 'Create', undefined, `${f.LicenceCode} ${f.Title}`);
    return res.data.Id;
  }

  public async updateLicence(id: number, f: ILicenceInput, loadedModified: string, extra?: any): Promise<void> {
    const cur = await this.getLicence(id);
    if (cur.Modified !== loadedModified) { throw new Error('This licence was changed by someone else while you were editing (V14). Reload and retry.'); }
    await this.checkLicenceCode(f.LicenceCode, id);
    const open = (await this.getAllocations(id)).filter(a => a.AllocationStatus === 'Open').length;
    if (f.PurchasedSeats !== undefined && f.PurchasedSeats < open) {
      throw new Error(`Purchased seats cannot be reduced below the ${open} open allocation(s) (V10). Reclaim seats first.`);
    }
    await this.items(LISTS.Licences).getById(id).update({ ...this.licPayload(f), ...(extra || {}) });
    await this.log('Licence', id, extra && extra.LastRenewedOn ? 'Renew' : 'Update', undefined,
      this.diff(cur, f, ['Title', 'LicenceCode', 'Vendor', 'LicenceModel', 'PurchasedSeats', 'StartDate', 'RenewalDate', 'Renewable', 'AutoRenew', 'BillingFrequency', 'Amount', 'CurrencyCode', 'LicenceStatus']));
  }

  /** Extend renewal: updates RenewalDate + LastRenewedOn only (never claims payment happened). */
  public async renewLicence(id: number, newDate: Date): Promise<void> {
    const l = await this.getLicence(id);
    if (l.StartDate && !sameOrAfter(newDate, l.StartDate)) { throw new Error('Renewal date cannot be before the start date (V04).'); }
    await this.updateLicence(id, { ...l, RenewalDate: newDate, Renewable: true }, l.Modified, { LastRenewedOn: toDateOnly(today()) });
  }

  public async cancelLicence(id: number): Promise<void> {
    const l = await this.getLicence(id);
    await this.items(LISTS.Licences).getById(id).update({ LicenceStatus: 'Cancelled' });
    await this.log('Licence', id, 'Cancel', l.LicenceStatus, 'Cancelled');
  }

  public async getAllocations(licenceId?: number): Promise<ILicenceAllocation[]> {
    let q = this.items(LISTS.LicenceAssignments).select(...LA_SEL).expand('TargetPerson');
    if (licenceId) { q = q.filter(`LicenceID eq ${licenceId}`); }
    return (await q.getAll()).map(mapLA);
  }
  public async getOpenAllocations(): Promise<ILicenceAllocation[]> {
    const r: any[] = await this.items(LISTS.LicenceAssignments).select(...LA_SEL).expand('TargetPerson').filter(`AllocationStatus eq 'Open'`).getAll();
    return r.map(mapLA);
  }

  /** WF04 – allocate seat. Pre-check, write, then post-write verification to resolve simultaneous allocations. */
  public async allocateSeat(licenceId: number, target: { person?: IPerson; assetId?: number; assetTag?: string }, date: Date, notes?: string): Promise<void> {
    const lic = await this.getLicence(licenceId);
    if (lic.LicenceStatus !== 'Active') { throw new Error(`New allocations are not allowed while the licence is ${lic.LicenceStatus}.`); }
    if (!isSeatManaged(lic)) { throw new Error('This licence model does not support seat allocation.'); }
    const seats = lic.PurchasedSeats || 0;
    const open = (await this.getAllocations(licenceId)).filter(a => a.AllocationStatus === 'Open');
    if (open.length >= seats) { throw new Error(`Seat limit reached (V10): ${open.length} of ${seats} seats are allocated.`); }
    const dup = open.some(a => (target.person && a.TargetPerson && a.TargetPerson.Id === target.person.Id) || (target.assetId && a.TargetAssetID === target.assetId));
    if (dup) { throw new Error('This target already holds an active seat on this licence (V09).'); }
    const label = target.person ? target.person.Title : target.assetTag;
    const created: any = await this.items(LISTS.LicenceAssignments).add({
      Title: `${lic.LicenceCode} - ${label} - ${toDateOnly(date)}`, LicenceID: licenceId, LicenceCodeSnapshot: lic.LicenceCode,
      TargetPersonId: target.person ? target.person.Id : null, TargetAssetID: target.assetId || null,
      AssignedDate: toDateOnly(date), AllocationStatus: 'Open', Notes: notes || null
    });
    const newId: number = created.data.Id;
    const after = (await this.getAllocations(licenceId)).filter(a => a.AllocationStatus === 'Open').sort((a, b) => a.Id - b.Id);
    const winners = after.slice(0, seats).map(a => a.Id);
    const samePerson = after.filter(a => a.Id !== newId && ((target.person && a.TargetPerson && a.TargetPerson.Id === target.person.Id) || (target.assetId && a.TargetAssetID === target.assetId)));
    if (winners.indexOf(newId) < 0 || samePerson.some(a => a.Id < newId)) {
      await this.items(LISTS.LicenceAssignments).getById(newId).update({
        AllocationStatus: 'Reclaimed', ReclaimedDate: toDateOnly(date), Notes: 'Auto-cancelled: concurrent allocation conflict'
      });
      await this.log('Licence Allocation', licenceId, 'Allocate', undefined, `Rejected: concurrent conflict for ${label}`, 'Failure');
      throw new Error('Another allocation was saved at the same moment and the seat limit/duplicate rule would be broken (V14). Your allocation was cancelled; refresh and retry.');
    }
    await this.log('Licence Allocation', licenceId, 'Allocate', undefined, `${lic.LicenceCode} → ${label}`);
  }

  public async reclaimSeat(allocationId: number, date: Date): Promise<void> {
    const r: any = await this.items(LISTS.LicenceAssignments).getById(allocationId).select(...LA_SEL).expand('TargetPerson').get();
    const a = mapLA(r);
    if (a.AllocationStatus !== 'Open') { throw new Error('This allocation is already reclaimed.'); }
    if (a.AssignedDate && !sameOrAfter(date, a.AssignedDate)) { throw new Error(`Reclaim date cannot precede the allocation date ${fmt(a.AssignedDate)} (V08).`); }
    await this.items(LISTS.LicenceAssignments).getById(allocationId).update({ AllocationStatus: 'Reclaimed', ReclaimedDate: toDateOnly(date) });
    await this.log('Licence Allocation', a.LicenceID, 'Reclaim', undefined, `${a.LicenceCodeSnapshot} seat ${allocationId} reclaimed`);
  }

  // ---------------------------------------------------------------- access requests (S11)
  public async getAccessRequests(): Promise<IAccessRequest[]> {
    const r: any[] = await this.items(LISTS.AccessRequests).select(...AR_SEL).expand('RequestedBy', 'ApprovedBy', 'ProvisionedBy').orderBy('Id', false).getAll();
    return r.map(x => ({ ...x, RequestedBy: P(x.RequestedBy), ApprovedBy: P(x.ApprovedBy), ProvisionedBy: P(x.ProvisionedBy) } as IAccessRequest));
  }

  public async addAccessRequest(email: string, role: string, notes?: string): Promise<void> {
    const e = email.trim().toLowerCase();
    const ex: any[] = await this.items(LISTS.AccessRequests).select('Id', 'AccessStatus').filter(`UserEmail eq '${esc(e)}'`).get();
    if (ex.some(x => ['Pending', 'Approved', 'Active'].indexOf(x.AccessStatus) >= 0)) { throw new Error('An open or active request already exists for this email.'); }
    const res: any = await this.items(LISTS.AccessRequests).add({
      Title: e, UserEmail: e, RequestedRole: role, AccessStatus: 'Pending', RequestedById: this.me.Id,
      RequestedOn: new Date().toISOString(), Notes: notes || null
    });
    await this.log('Access', res.data.Id, 'Request', undefined, `${e} as ${role}`);
  }

  public async setAccessStatus(req: IAccessRequest, status: string, all: IAccessRequest[]): Promise<void> {
    if (status === 'Approved' && req.UserEmail === (this.me.EMail || '').toLowerCase()) { throw new Error('You cannot approve your own access request.'); }
    if (status === 'Revoked' && req.RequestedRole === 'Admin' && req.AccessStatus === 'Active') {
      const otherAdmins = all.filter(r => r.Id !== req.Id && r.RequestedRole === 'Admin' && r.AccessStatus === 'Active');
      if (!otherAdmins.length) { throw new Error('Cannot revoke the last active Owner/Admin (V15).'); }
    }
    const now = new Date().toISOString();
    const upd: any = { AccessStatus: status };
    if (status === 'Approved' || status === 'Rejected') { upd.ApprovedById = this.me.Id; upd.ApprovedOn = now; }
    if (status === 'Active') { upd.ProvisionedById = this.me.Id; upd.ProvisionedOn = now; }
    await this.items(LISTS.AccessRequests).getById(req.Id).update(upd);
    await this.log('Access', req.Id, status, req.AccessStatus, `${req.UserEmail} (${req.RequestedRole})`);
  }

  // ---------------------------------------------------------------- settings / reminders / reconciliation
  public async getSettings(): Promise<{ [k: string]: string }> {
    const r: any[] = await this.items(LISTS.Settings).select('Title', 'SettingValue').getAll();
    const o: { [k: string]: string } = {};
    r.forEach(x => { o[x.Title] = x.SettingValue; });
    return o;
  }

  public async saveSetting(key: string, value: string): Promise<void> {
    const ex: any[] = await this.items(LISTS.Settings).select('Id', 'SettingValue').filter(`Title eq '${esc(key)}'`).top(1).get();
    if (ex.length) { await this.items(LISTS.Settings).getById(ex[0].Id).update({ SettingValue: value }); }
    else { await this.items(LISTS.Settings).add({ Title: key, SettingValue: value }); }
    await this.log('Settings', ex.length ? ex[0].Id : undefined, 'Update', ex.length ? ex[0].SettingValue : undefined, `${key} = ${value}`);
  }

  public async getFailedReminders(): Promise<any[]> {
    return this.items(LISTS.ReminderLog).select('Id', 'Title', 'EntityType', 'EntityID', 'ExpiryDate', 'ThresholdDays', 'SendStatus', 'ErrorMessage')
      .filter(`SendStatus eq 'Failed'`).orderBy('Id', false).top(50).get();
  }

  /** Flags inconsistent state (spec s7: lists + flows are not ACID). */
  public async reconcile(): Promise<string[]> {
    const issues: string[] = [];
    const [assets, open, lics, allocs] = await Promise.all([
      this.getAssets(),
      this.items(LISTS.AssetAssignments).select('Id', 'AssetID', 'AssetTagSnapshot').filter(`AssignmentStatus eq 'Open'`).getAll() as Promise<any[]>,
      this.getLicences(), this.getOpenAllocations()
    ]);
    const openBy: { [k: number]: any[] } = {};
    open.forEach(o => { (openBy[o.AssetID] = openBy[o.AssetID] || []).push(o); });
    assets.forEach(a => {
      const o = openBy[a.Id] || [];
      if (a.AssetStatus === 'Assigned' && !o.length) { issues.push(`${a.AssetTag}: status Assigned but no open assignment.`); }
      if (a.AssetStatus !== 'Assigned' && o.length) { issues.push(`${a.AssetTag}: status ${a.AssetStatus} but has an open assignment (#${o[0].Id}).`); }
      if (o.length > 1) { issues.push(`${a.AssetTag}: ${o.length} open assignments (only one allowed).`); }
      if (a.AssetStatus === 'Assigned' && o.length === 1 && a.CurrentAssignmentID !== o[0].Id) { issues.push(`${a.AssetTag}: CurrentAssignmentID does not match the open assignment.`); }
    });
    lics.forEach(l => {
      const n = allocs.filter(x => x.LicenceID === l.Id).length;
      if (isSeatManaged(l) && n > (l.PurchasedSeats || 0)) { issues.push(`${l.LicenceCode}: ${n} open allocations exceed ${l.PurchasedSeats} purchased seats.`); }
      if (!isSeatManaged(l) && n > 0) { issues.push(`${l.LicenceCode}: has open allocations but the model is not seat-managed.`); }
    });
    return issues;
  }

  // ---------------------------------------------------------------- documents (restricted library)
  public async getDocs(type: 'Asset' | 'Licence', id: number): Promise<IDocument[]> {
    return this.items(LISTS.Documents).select('Id', 'FileLeafRef', 'FileRef', 'DocumentType')
      .filter(`RelatedEntityType eq '${type}' and RelatedItemID eq ${id}`).get();
  }

  public async uploadDoc(type: 'Asset' | 'Licence', id: number, docType: string, file: File): Promise<void> {
    const lib = sp.web.lists.getByTitle(LISTS.Documents);
    const res: any = await lib.rootFolder.files.add(`${type}-${id}-${Date.now()}-${file.name}`, file, true);
    const item: any = await res.file.getItem();
    await item.update({ RelatedEntityType: type, RelatedItemID: id, DocumentType: docType });
    await this.log(type, id, 'Document', undefined, `${docType}: ${file.name}`);
  }
}

export const DataService = new DataServiceImpl();
