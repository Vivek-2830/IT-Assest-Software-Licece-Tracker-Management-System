import * as React from 'react';
import {
  DetailsList, DetailsListLayoutMode, IColumn, SelectionMode,
  SearchBox, Dropdown, IDropdownOption, PrimaryButton, DefaultButton,
  IconButton, MessageBar, MessageBarType, Spinner, SpinnerSize,
  Dialog, DialogType, DialogFooter, TextField,
  IPersonaProps,
  NormalPeoplePicker
} from '@fluentui/react';
import { AssetService } from '../../services/AssetService';
import { IAsset, IAssetFormValues } from '../../models/IAsset';
import styles from './Assets.module.scss';
import { sp } from '@pnp/sp/rest';

const emptyForm: IAssetFormValues = {
  Title: '', AssetTag: '', AssetType: 'Laptop', Manufacturer: '', Model: '',
  SerialNumber: '', PurchaseDate: '', PurchaseCost: '', Supplier: '',
  InvoiceReference: '', WarrantyEnd: '', OfficeLocation: 'Unspecified',
  HolderEmail: '', Notes: ''
};

const assetTypeOptions: IDropdownOption[] = [
  'Laptop', 'Desktop', 'Monitor', 'Phone', 'Tablet', 'Network', 'Accessory', 'Other'
].map((v: string) => ({ key: v, text: v }));

const locationOptions: IDropdownOption[] = [
  'UK', 'India', 'Other', 'Unspecified'
].map((v: string) => ({ key: v, text: v }));

const statusOptions: IDropdownOption[] = [
  { key: 'All', text: 'All statuses' },
  { key: 'Available', text: 'Available' },
  { key: 'Assigned', text: 'Assigned' },
  { key: 'Repair', text: 'Repair' },
  { key: 'Retired', text: 'Retired' }
];

function dateForInput(value?: string): string {
  if (!value) { return ''; }
  const date: Date = new Date(value);
  return isNaN(date.getTime()) ? '' : date.toISOString().substring(0, 10);
}

function mapToForm(asset?: IAsset): IAssetFormValues {
  if (!asset) { return { ...emptyForm }; }
  return {
    Title: asset.Title || '',
    AssetTag: asset.AssetTag || '',
    AssetType: asset.AssetType || 'Laptop',
    Manufacturer: asset.Manufacturer || '',
    Model: asset.Model || '',
    SerialNumber: asset.SerialNumber || '',
    PurchaseDate: dateForInput(asset.PurchaseDate),
    PurchaseCost: asset.PurchaseCost === null || asset.PurchaseCost === undefined ? '' : String(asset.PurchaseCost),
    Supplier: asset.Supplier || '',
    InvoiceReference: asset.InvoiceReference || '',
    WarrantyEnd: dateForInput(asset.WarrantyEnd),
    OfficeLocation: asset.OfficeLocation || 'Unspecified',
    HolderEmail: asset.CurrentHolder ? (asset.CurrentHolder.EMail || '') : '',
    Notes: asset.Notes || ''
  };
}

export const Assets: React.FC = () => {
  const [items, setItems] = React.useState<IAsset[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [saving, setSaving] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const [search, setSearch] = React.useState<string>('');
  const [status, setStatus] = React.useState<string>('All');
  const [dialogOpen, setDialogOpen] = React.useState<boolean>(false);
  const [editing, setEditing] = React.useState<IAsset | undefined>(undefined);
  const [form, setForm] = React.useState<IAssetFormValues>({ ...emptyForm });
  const [selectedHolder, setSelectedHolder] = React.useState<IPersonaProps[]>([]);

  const loadAssets = React.useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      setItems(await AssetService.getAssets(false));
    } catch (e) {
      setError(e && e.message ? e.message : 'Unable to load IT_Assets. Verify the list and internal field names.');
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => { loadAssets(); }, [loadAssets]);

  const openNew = (): void => {
    setEditing(undefined);
    setForm({ ...emptyForm });
    setSelectedHolder([]);
    setDialogOpen(true);
  };

  const openEdit = (asset: IAsset): void => {
    setEditing(asset);
    setForm(mapToForm(asset));

    if (asset.CurrentHolder) {
      setSelectedHolder([{
        // key: asset.CurrentHolder.EMail,
        text: asset.CurrentHolder.Title,
        secondaryText: asset.CurrentHolder.EMail
      }]);
    } else {
      setSelectedHolder([]);
    }

    setDialogOpen(true);
  };
  const setField = (field: keyof IAssetFormValues, value: string): void => {
    setForm((old: IAssetFormValues) => ({ ...old, [field]: value }));
  };

  const save = async (): Promise<void> => {
    setError('');

    if (!form.Title.trim() || !form.AssetTag.trim()) {
      setError('Asset Name and Asset Tag are required.');
      return;
    }

    if (
      form.PurchaseCost &&
      (isNaN(Number(form.PurchaseCost)) || Number(form.PurchaseCost) < 0)
    ) {
      setError('Purchase Cost must be a valid number greater than or equal to zero.');
      return;
    }

    if (
      form.PurchaseDate &&
      form.WarrantyEnd &&
      new Date(form.WarrantyEnd) < new Date(form.PurchaseDate)
    ) {
      setError('Warranty End cannot be earlier than Purchase Date.');
      return;
    }

    setSaving(true);

    try {
      let holderEmail = '';

      // Resolve the selected People Picker user
      if (selectedHolder.length > 0) {
        holderEmail =
          selectedHolder[0].secondaryText ||
          String(selectedHolder[0] || '');

        if (holderEmail) {
          const ensuredUser = await sp.web.ensureUser(holderEmail);

          holderEmail = ensuredUser.data.Email || holderEmail;
        }
      }

      // Pass selected user's email to AssetService
      const assetForm: IAssetFormValues = {
        ...form,
        HolderEmail: holderEmail
      };

      if (editing) {
        await AssetService.updateAsset(editing.Id, assetForm);
      } else {
        await AssetService.createAsset(assetForm);
      }

      setDialogOpen(false);
      setSelectedHolder([]);

      await loadAssets();

    } catch (e) {
      setError(
        e && e.message
          ? e.message
          : 'Unable to save asset.'
      );
    } finally {
      setSaving(false);
    }
  };

  const archive = async (asset: IAsset): Promise<void> => {
    if (!window.confirm('Archive ' + asset.AssetTag + '? The record will be retained and hidden from the active list.')) {
      return;
    }
    try {
      await AssetService.archiveAsset(asset.Id);
      await loadAssets();
    } catch (e) {
      setError(e && e.message ? e.message : 'Unable to archive asset.');
    }
  };

  const onResolveSuggestions = async (
    filterText: string
  ): Promise<IPersonaProps[]> => {
    if (!filterText || filterText.length < 2) {
      return [];
    }

    const searchText = filterText.replace(/'/g, "''");

    const users = await sp.web.siteUsers
      .filter(
        `substringof('${searchText}',Title) or substringof('${searchText}',Email)`
      )
      .top(10)
      .get();

    return users.map((user: any) => ({
      key: user.Email || user.LoginName,
      text: user.Title,
      secondaryText: user.Email,
      id: user.Id
    }));
  };

  const filtered: IAsset[] = items.filter((item: IAsset) => {
    const term: string = search.toLowerCase();
    const holder: string = item.CurrentHolder && item.CurrentHolder.Title ? item.CurrentHolder.Title.toLowerCase() : '';
    const matchesSearch: boolean = !term ||
      (item.Title || '').toLowerCase().indexOf(term) !== -1 ||
      (item.AssetTag || '').toLowerCase().indexOf(term) !== -1 ||
      (item.SerialNumber || '').toLowerCase().indexOf(term) !== -1 ||
      holder.indexOf(term) !== -1;
    return matchesSearch && (status === 'All' || item.AssetStatus === status);
  });

  const columns: IColumn[] = [
    { key: 'tag', name: 'Asset Tag', minWidth: 90, maxWidth: 120, onRender: (item: IAsset) => <span className={styles.assetTag}>{item.AssetTag}</span> },
    { key: 'name', name: 'Asset Name', fieldName: 'Title', minWidth: 120, isResizable: true },
    { key: 'type', name: 'Type', fieldName: 'AssetType', minWidth: 80, maxWidth: 110 },
    { key: 'model', name: 'Model', fieldName: 'Model', minWidth: 90, isResizable: true },
    { key: 'holder', name: 'Current Holder', minWidth: 120, onRender: (item: IAsset) => item.CurrentHolder ? item.CurrentHolder.Title : '—' },
    { key: 'status', name: 'Status', minWidth: 90, onRender: (item: IAsset) => <span className={styles['status' + item.AssetStatus]}>{item.AssetStatus}</span> },
    { key: 'warranty', name: 'Warranty End', minWidth: 100, onRender: (item: IAsset) => item.WarrantyEnd ? new Date(item.WarrantyEnd).toLocaleDateString() : '—' },
    {
      key: 'actions', name: 'Actions', minWidth: 75, maxWidth: 90,
      onRender: (item: IAsset) => (
        <div className={styles.actions}>
          <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit asset" onClick={() => openEdit(item)} />
          <IconButton iconProps={{ iconName: 'Archive' }} title="Archive" ariaLabel="Archive asset" onClick={() => archive(item)} />
        </div>
      )
    }
  ];

  return (
    <div className={styles.assets}>
      <div className={styles.toolbar}>
        <div><h2>IT Assets</h2><p>Register, search and maintain your organisation's IT equipment.</p></div>
        <PrimaryButton text="Add Asset" iconProps={{ iconName: 'Add' }} onClick={openNew} />
      </div>

      {error && <MessageBar messageBarType={MessageBarType.error} isMultiline onDismiss={() => setError('')}>{error}</MessageBar>}

      <div className={styles.filters}>
        <SearchBox placeholder="Search name, asset tag, serial number or holder" value={search} onChange={(_, value) => setSearch(value || '')} />
        <Dropdown selectedKey={status} options={statusOptions} onChange={(_, option) => setStatus(String(option.key))} styles={{ root: { width: 170 } }} />
        <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={loadAssets} />
      </div>

      <div className={styles.table}>
        {loading ? <Spinner size={SpinnerSize.large} label="Loading assets..." /> : (
          <DetailsList
            items={filtered}
            columns={columns}
            setKey="assets"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            isHeaderVisible={true}
          />
        )}
        {!loading && filtered.length === 0 && <div className={styles.empty}>No assets found. Add an asset or change your filters.</div>}
      </div>

      <Dialog
        hidden={!dialogOpen}
        onDismiss={() => setDialogOpen(false)}
        dialogContentProps={{ type: DialogType.largeHeader, title: editing ? 'Edit Asset' : 'Add New Asset', subText: 'Enter the asset information below.' }}
        minWidth={640}
        modalProps={{ isBlocking: saving }}
      >
        <div className={styles.formGrid}>
          <TextField label="Asset Name" required value={form.Title} onChange={(_, v) => setField('Title', v || '')} />
          <TextField label="Asset Tag" required value={form.AssetTag} onChange={(_, v) => setField('AssetTag', v || '')} />
          <Dropdown label="Asset Type" selectedKey={form.AssetType} options={assetTypeOptions} onChange={(_, o) => setField('AssetType', String(o.key))} />
          <TextField label="Manufacturer" value={form.Manufacturer} onChange={(_, v) => setField('Manufacturer', v || '')} />
          <TextField label="Model" value={form.Model} onChange={(_, v) => setField('Model', v || '')} />
          <TextField label="Serial Number" value={form.SerialNumber} onChange={(_, v) => setField('SerialNumber', v || '')} />
          <TextField label="Purchase Date" type="date" value={form.PurchaseDate} onChange={(_, v) => setField('PurchaseDate', v || '')} />
          <TextField label="Purchase Cost" type="number" prefix="£" value={form.PurchaseCost} onChange={(_, v) => setField('PurchaseCost', v || '')} />
          <TextField label="Supplier" value={form.Supplier} onChange={(_, v) => setField('Supplier', v || '')} />
          <TextField label="Invoice Reference" value={form.InvoiceReference} onChange={(_, v) => setField('InvoiceReference', v || '')} />
          <TextField label="Warranty End" type="date" value={form.WarrantyEnd} onChange={(_, v) => setField('WarrantyEnd', v || '')} />
          <Dropdown label="Office Location" selectedKey={form.OfficeLocation} options={locationOptions} onChange={(_, o) => setField('OfficeLocation', String(o.key))} />
          {/* <TextField label="Current Holder Email" type="email" value={form.HolderEmail} onChange={(_, v) => setField('HolderEmail', v || '')} /> */}
          <div className="formField">
            <label>Current Holder</label>

            <NormalPeoplePicker
              onResolveSuggestions={onResolveSuggestions}
              onChange={(items) => {
                setSelectedHolder(items || []);
              }}
              selectedItems={selectedHolder}
              itemLimit={1}
              pickerSuggestionsProps={{
                suggestionsHeaderText: 'Suggested Users',
                noResultsFoundText: 'No users found'
              }}
              inputProps={{
                placeholder: 'Search employee by name or email...'
              }}
            />
          </div>
          <div className={styles.fullWidth}><TextField label="Notes" multiline rows={3} value={form.Notes} onChange={(_, v) => setField('Notes', v || '')} /></div>
        </div>
        <DialogFooter>
          <PrimaryButton text={saving ? 'Saving...' : 'Save'} onClick={save} disabled={saving} />
          <DefaultButton text="Cancel" onClick={() => setDialogOpen(false)} disabled={saving} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
