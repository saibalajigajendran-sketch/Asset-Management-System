import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { Icon, initializeIcons } from '@fluentui/react';
import styles from './AssetManagement.module.scss';
import type { IAssetManagementProps } from './IAssetManagementProps';
import { deleteSharePointListItem, getEntraEmployees, getSharePointListFields, getSharePointListItems, getSharePointUserNames, saveSharePointListItem, type IEntraEmployee, type ISharePointField, type SharePointUserNameMap } from './hardwareLaptopSharePointService';
import AssetAttachments, { type AssetAttachmentsHandle } from './AssetAttachments';
import { assetRegisterUrls, dashboardRegisters } from './dashboardConfig';
import { orderAssetFormFields } from './assetFormFieldOrder';
import { validateAssetForm } from './assetFormValidation';
import { createExcelWorkbook } from './assetExcelExport';
import Dashboard from './Dashboard';
import LaptopViewTabs from './LaptopViewTabs';

initializeIcons();

type AssetStatus = 'Active' | 'In Maintenance' | 'In Stock' | 'Allocated' | 'Retired';
type Asset = {
  id: string;
  name: string;
  status: AssetStatus;
  assignedTo: string;
  purchaseDate: string;
  value: number;
  location: string;
  [key: string]: string | number;
};
type Category = {
  key: string;
  name: string;
  icon: string;
  description: string;
  columns: { key: string; label: string }[];
  fields: string[];
  incidents: string[];
  assets: Asset[];
};

const getWarrantyStatusValues = (asset: Asset): string[] => Object.keys(asset)
    .filter(key => key.toLowerCase().replace(/[^a-z0-9]/g, '').endsWith('warrantystatus') && asset[key] !== undefined && asset[key] !== null)
    .map(key => String(asset[key]).trim())
    .filter(Boolean);
const getCategoryPrimaryStatus = (asset: Asset, categoryKey: string): string => {
  const categoryStatus = categoryKey === 'outsourced-services' ? String(asset.vendorStatus || '').trim() : '';
  return categoryStatus || String(asset.status || '').trim();
};
const getAssetFilterStatuses = (asset: Asset, categoryKey: string): string[] => [getCategoryPrimaryStatus(asset, categoryKey), ...getWarrantyStatusValues(asset)].filter(Boolean);
const formatDateOnly = (value: string): string => /^\d{4}-\d{2}-\d{2}T/.test(value) ? value.slice(0, 10) : value;
const getAssetFieldValue = (asset: Asset, field: ISharePointField): string => {
  const normalizeKey = (key: string): string => key.toLowerCase().replace(/[^a-z0-9]/g, '');
  const fieldKey = Object.keys(asset).find(key => normalizeKey(key) === normalizeKey(field.InternalName))
    || Object.keys(asset).find(key => normalizeKey(key) === normalizeKey(field.Title));
  const value = fieldKey ? asset[fieldKey] : undefined;
  return value === undefined || value === null || String(value).trim() === '' ? 'N/A' : formatDateOnly(String(value));
};

const ciaFields = ['Confidentiality', 'Integrity', 'Availability'];
const createCategory = (
  key: string,
  name: string,
  icon: string,
  description: string,
  columns: [string, string, string, string],
  fields: string[],
  incidents: string[],
  assets: Asset[]
): Category => ({
  key,
  name,
  icon,
  description,
  columns: [{ key: columns[0], label: columns[1] }, { key: columns[2], label: columns[3] }],
  fields: [...fields, ...ciaFields],
  incidents,
  assets
});

const initialCategories: Category[] = [
  createCategory('hardware-laptop', 'Hardware - Laptops, Desktops', 'LaptopSecure', 'Laptops and portable computing hardware.', ['serialNumber', 'SerialNumber', 'modelNo', 'ModelNo.'], ['Department', 'UserID', 'UserRole', 'AssetSpecification', 'OperatingSystemDetails', 'WarrantyStatus', 'ExpiryDateofSupport', 'AssetOwner', 'AssetCustodian'], ['Battery failure', 'Screen damage', 'Keyboard fault', 'Overheating', 'OS crash'], [
    { id: 'AST-LAP-101', name: 'MacBook Pro 16"', status: 'Active', assignedTo: 'Sarah Jenkins', purchaseDate: '2024-01-15', value: 2499, location: 'HQ - Floor 3', serialNumber: 'C02G901NMD6R', modelNo: 'MacBook Pro 16-inch', extra_asset_specifications: 'M2 Max / 32GB' },
    { id: 'AST-LAP-102', name: 'Lenovo ThinkPad X1', status: 'In Maintenance', assignedTo: 'David Miller', purchaseDate: '2022-05-20', value: 1350, location: 'Remote - NY', serialNumber: 'PF3K8912', modelNo: 'ThinkPad X1', extra_asset_specifications: 'i7 / 16GB' },
    { id: 'AST-LAP-103', name: 'Dell Latitude 7440', status: 'In Stock', assignedTo: 'Unassigned', purchaseDate: '2024-02-10', value: 1499, location: 'IT Storage Room', serialNumber: 'DL-7440-221', modelNo: 'Latitude 7440', extra_asset_specifications: 'i5 / 16GB' }
  ]),
  createCategory('network-devices', 'Network Devices', 'NetworkTower', 'Routers, switches, firewalls and access points.', ['ipAddress', 'IP address', 'firmware', 'Firmware'], ['Asset Owner', 'Asset Custodian', 'Model No.', 'Warranty Status', 'Support Expiry Date', 'Specifications'], ['Link down', 'High latency', 'Firmware bug', 'Port failure'], [
    { id: 'AST-NET-201', name: 'Cisco Catalyst 9300', status: 'Active', assignedTo: 'Network Team', purchaseDate: '2022-08-12', value: 5500, location: 'Server Room B', ipAddress: '192.168.1.1', firmware: 'IOS-XE 17.3' },
    { id: 'AST-NET-202', name: 'Palo Alto PA-440', status: 'Active', assignedTo: 'Security Ops', purchaseDate: '2023-01-20', value: 3200, location: 'Server Room A', ipAddress: '10.0.0.1', firmware: 'PAN-OS 10.2' },
    { id: 'AST-NET-203', name: 'Ubiquiti UniFi 6 Pro', status: 'In Maintenance', assignedTo: 'Facilities', purchaseDate: '2023-06-15', value: 159, location: 'HQ - Ceiling F4', ipAddress: '192.168.10.45', firmware: 'v6.5.28' }
  ]),
  createCategory('asset-tracker', 'Asset tracker', 'ClipboardList', 'Track assigned assets, ownership and movement.', ['owner', 'Current owner', 'tag', 'Asset tag'], ['AssetDescription', 'DateOfMovement', 'AssetMovedFrom', 'AssetMovedTo', 'AssetStatus', 'AssignedTo'], ['Asset not found', 'Wrong assignment', 'Tag damaged'], [
    { id: 'AST-TRK-301', name: 'Executive Laptop Bag', status: 'Active', assignedTo: 'CEO Office', purchaseDate: '2021-04-10', value: 195, location: 'Executive Suite', owner: 'CEO Office', tag: 'TAG-0301' },
    { id: 'AST-TRK-302', name: 'Docking Station Kit', status: 'Active', assignedTo: 'Jane Smith', purchaseDate: '2022-09-01', value: 280, location: 'HQ - Floor 3', owner: 'Jane Smith', tag: 'TAG-0302' }
  ]),
  createCategory('outsourced-services', 'Outsourced Services', 'BusinessCenterLogo', 'Vendor-managed services and contracts.', ['vendor', 'Vendor', 'contractEnd', 'Contract end'], ['Vendor ID', 'Vendor Status', 'Vendor Name', 'Purpose'], ['SLA breach', 'Billing dispute', 'Service outage'], [
    { id: 'AST-OUT-401', name: 'Managed Print Service', status: 'Active', assignedTo: 'Operations', purchaseDate: '2023-04-15', value: 12000, location: 'All Floors', vendor: 'Xerox Managed', contractEnd: '2027-04-14' },
    { id: 'AST-OUT-402', name: '24x7 Network Monitoring', status: 'In Maintenance', assignedTo: 'IT Infrastructure', purchaseDate: '2022-10-05', value: 18500, location: 'Remote', vendor: 'NetWatch Ltd', contractEnd: '2026-12-31' }
  ]),
  createCategory('software', 'Software', 'Permissions', 'SaaS subscriptions, keys and developer licenses.', ['licenseKey', 'License key', 'expiryDate', 'Expiry'], ['Version', 'Software Platform', 'License Type', 'License Expiry Date', 'Quantity Purchased', 'Quantity Used'], ['Activation error', 'Seat shortage', 'Renewal issue'], [
    { id: 'AST-SFT-501', name: 'JetBrains All Products Pack', status: 'Active', assignedTo: 'Engineering Team', purchaseDate: '2025-01-01', value: 649, location: 'Cloud Subscription', licenseKey: 'JB-ENT-994821', expiryDate: '2027-01-01' },
    { id: 'AST-SFT-502', name: 'Adobe Creative Cloud', status: 'Active', assignedTo: 'Design Team', purchaseDate: '2025-03-10', value: 1200, location: 'Cloud Subscription', licenseKey: 'ADOBE-CC-8812', expiryDate: '2026-03-10' }
  ]),
  createCategory('hardware-server', 'Hardware - Server', 'Server', 'Server hardware, specifications and service details.', ['serialNumber', 'Serial number', 'specs', 'Processor / RAM'], ['Asset Owner', 'Asset Custodian', 'Model No.', 'Warranty Status', 'Support Expiry Date', 'Purpose', 'Host Name', 'Operating System'], ['Disk degraded', 'Fan failure', 'Unplanned reboot', 'PSU fault'], [
    { id: 'AST-SRV-601', name: 'PowerEdge R750', status: 'Active', assignedTo: 'IT Infrastructure', purchaseDate: '2023-03-01', value: 8900, location: 'Server Room A', serialNumber: 'SRV-DELL-992', specs: 'Dual Xeon / 128GB' },
    { id: 'AST-SRV-602', name: 'HPE ProLiant DL380', status: 'In Maintenance', assignedTo: 'IT Infrastructure', purchaseDate: '2022-07-18', value: 7600, location: 'Server Room B', serialNumber: 'SRV-HPE-4410', specs: 'Xeon Gold / 256GB' }
  ]),
  createCategory('cloud-provider', 'Cloud Providers', 'Cloud', 'Cloud services, ownership and contract details.', ['provider', 'Provider', 'plan', 'Plan / region'], ['Cloud Service Partner', 'Vendor ID', 'Service Type', 'Purpose', 'Business Owner', 'Technical Owner', 'Agreement Expiry Date'], ['Billing anomaly', 'Access outage', 'Quota limit'], [
    { id: 'AST-CLD-701', name: 'Azure Enterprise Subscription', status: 'Active', assignedTo: 'Cloud Team', purchaseDate: '2023-02-10', value: 4500, location: 'Central India', provider: 'Microsoft Azure', plan: 'Enterprise / South India' },
    { id: 'AST-CLD-702', name: 'AWS Production Account', status: 'Active', assignedTo: 'DevOps', purchaseDate: '2022-11-01', value: 6200, location: 'ap-south-1', provider: 'Amazon AWS', plan: 'Business / Mumbai' }
  ]),
  createCategory('hardware-accessories', 'Hardware - Accessories', 'KeyboardClassic', 'Peripherals and hardware accessories.', ['accType', 'Type', 'connection', 'Connection'], ['Asset Owner', 'Asset Custodian', 'Model No.', 'Warranty Status', 'Support Expiry Date', 'Total Count'], ['Not detected', 'Physical damage', 'Cable fault'], [
    { id: 'AST-ACC-801', name: 'Logitech MX Master 3', status: 'In Stock', assignedTo: 'Unassigned', purchaseDate: '2024-02-10', value: 99, location: 'IT Storage Room', accType: 'Mouse', connection: 'Bluetooth' },
    { id: 'AST-ACC-802', name: 'Jabra Evolve2 75', status: 'Active', assignedTo: 'Priya Raman', purchaseDate: '2023-05-12', value: 349, location: 'HQ - Floor 2', accType: 'Headset', connection: 'USB-C' },
    { id: 'AST-ACC-803', name: 'Dell WD19 Dock', status: 'Active', assignedTo: 'Karthik Iyer', purchaseDate: '2023-08-02', value: 210, location: 'HQ - Floor 3', accType: 'Dock', connection: 'USB-C' }
  ]),
  createCategory('intellectual-property', ' Intellectual Property', 'Lightbulb', 'Registrations and compliance details.', ['regNumber', 'Registration no.', 'renewalDue', 'Renewal due'], ['Asset Owner', 'Asset Custodian', 'Asset Description', 'Date of Renewal'], ['Renewal overdue', 'Filing query', 'Infringement alert'], [
    { id: 'AST-IPR-901', name: 'Trademark "Asset Management"', status: 'Active', assignedTo: 'Legal', purchaseDate: '2021-07-01', value: 15000, location: 'Legal Archive', regNumber: 'TM-2021-44871', renewalDue: '2031-07-01' },
    { id: 'AST-IPR-902', name: 'Patent - Asset Tag Scanner', status: 'Active', assignedTo: 'R&D', purchaseDate: '2022-03-14', value: 42000, location: 'Legal Archive', regNumber: 'PT-2022-10934', renewalDue: '2027-03-14' }
  ]),
  createCategory('media-devices', 'Media Device', 'Video', 'Projectors, cameras, microphones and displays.', ['resolution', 'Resolution', 'room', 'Room'], ['Asset Owner', 'Asset Custodian', 'Model No.', 'Warranty Status', 'Support Expiry Date', 'Media Type', 'Capacity'], ['No audio', 'Display fault', 'Camera fault'], [
    { id: 'AST-MED-1001', name: 'Logitech Rally Bar', status: 'Active', assignedTo: 'Boardroom', purchaseDate: '2023-09-15', value: 3999, location: 'HQ - Floor 4', resolution: '4K Ultra HD', room: 'Boardroom A' },
    { id: 'AST-MED-1002', name: 'Samsung Flip 85"', status: 'In Maintenance', assignedTo: 'Training Team', purchaseDate: '2022-12-01', value: 4200, location: 'HQ - Floor 2', resolution: '4K Touch', room: 'Training Room' }
  ]),
  createCategory('mobile-devices', 'Mobile Devices', 'Phone', 'Smartphones, tablets and handheld scanners.', ['imeiNumber', 'IMEI / Serial', 'carrier', 'Carrier'], ['Asset Owner', 'Asset Custodian', 'Model No.', 'Warranty Status', 'Support Expiry Date', 'Current OS', 'Capacity'], ['Screen crack', 'Battery swelling', 'MDM sync failure'], [
    { id: 'AST-MOB-1101', name: 'iPhone 15 Pro 256GB', status: 'Active', assignedTo: 'Mark Johnson', purchaseDate: '2024-02-20', value: 1099, location: 'Field Operations', imeiNumber: '35698412093841', carrier: 'Verizon Enterprise' },
    { id: 'AST-MOB-1102', name: 'Galaxy Tab S9', status: 'In Stock', assignedTo: 'Unassigned', purchaseDate: '2024-04-02', value: 799, location: 'IT Storage Room', imeiNumber: 'R52T90X1KLM', carrier: 'Not activated' }
  ]),
  createCategory('client-accessor', 'Client Accessories', 'Contact', 'Client access equipment and facility assets.', ['clientName', 'Client', 'accessLevel', 'Access level'], ['Asset Owner', 'Asset Custodian', 'Model No.', 'Warranty Status', 'Support Expiry Date', 'Current OS', 'Total Count'], ['Access denied', 'Token expired', 'Device offline'], [
    { id: 'AST-CLA-1201', name: 'Client VPN Token - Acme', status: 'Active', assignedTo: 'Acme Corp', purchaseDate: '2024-01-08', value: 60, location: 'Client Site', clientName: 'Acme Corp', accessLevel: 'Read / Write' },
    { id: 'AST-CLA-1202', name: 'Secure Badge Reader Kit', status: 'Active', assignedTo: 'Globex Ltd', purchaseDate: '2023-10-19', value: 420, location: 'Client Site', clientName: 'Globex Ltd', accessLevel: 'Read only' }
  ])
];

const emptyCategories = initialCategories.map(category => ({ ...category, assets: [] as Asset[] }));

const getStatusClass = (status: string): string => {
  const normalizedStatus = status.trim().toLocaleLowerCase();
  if (normalizedStatus === 'active') return styles.statusActive;
  if (normalizedStatus.includes('maintenance') || normalizedStatus.includes('repair')) return styles.statusMaintenance;
  if (normalizedStatus.includes('retired') || normalizedStatus.includes('expired') || normalizedStatus.includes('disposed')) return styles.statusRetired;
  if (normalizedStatus.includes('stock') || normalizedStatus.includes('allocat') || normalizedStatus.includes('available')) return styles.statusStock;
  return styles.statusOther;
};
const getMetricClass = (status: string): string => {
  const normalizedStatus = status.trim().toLocaleLowerCase();
  if (normalizedStatus === 'active') return styles.metricActive;
  if (normalizedStatus.includes('maintenance') || normalizedStatus.includes('repair')) return styles.metricMaintenance;
  if (normalizedStatus.includes('retired') || normalizedStatus.includes('expired') || normalizedStatus.includes('disposed')) return styles.metricRetired;
  if (normalizedStatus.includes('stock') || normalizedStatus.includes('allocat') || normalizedStatus.includes('available')) return styles.metricStock;
  return styles.metricOther;
};
const renderAssetColumnValue = (asset: Asset, column: { key: string; label: string }): React.ReactNode => {
  const value = asset[column.key];
  const normalizedLabel = column.label.replace(/[^a-z]/gi, '').toLowerCase();
  if ((normalizedLabel === 'assetstatus' || normalizedLabel === 'warrantystatus' || normalizedLabel === 'vendorstatus') && value !== undefined && value !== '') {
    const normalizedValue = String(value).toLowerCase();
    const tone = normalizedValue === 'active'
      ? styles.statusActive
      : normalizedValue.indexOf('maintenance') >= 0
        ? styles.statusMaintenance
        : normalizedValue.indexOf('retired') >= 0 || normalizedValue.indexOf('expired') >= 0
          ? styles.statusRetired
          : styles.statusStock;
    return <span className={`${styles.status} ${tone}`}>{String(value)}</span>;
  }
  return value === undefined || value === null || value === '' ? 'N/A' : formatDateOnly(String(value));
};
const fieldId = (label: string): string => `extra_${label.toLowerCase().replace(/[^a-z0-9]+/g, '_')}`;
const normalizeFieldName = (label: string): string => label
  .replace(/_x([0-9a-f]{4})_/gi, (_match, hex: string) => String.fromCharCode(parseInt(hex, 16)))
  .toLowerCase()
  .replace(/[^a-z0-9]/g, '');
const humanizeLabel = (label: string): string => label
  .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
  .replace(/([A-Z])([A-Z][a-z])/g, '$1 $2')
  .replace(/dateof/gi, 'Date Of')
  .replace(/[_]+/g, ' ')
  .replace(/\s+/g, ' ')
  .trim();
const getRowValue = (row: Record<string, unknown>, names: string[], fields: ISharePointField[] = []): unknown => {
  for (let nameIndex = 0; nameIndex < names.length; nameIndex += 1) {
    const exactKey = Object.keys(row).filter(key => normalizeFieldName(key) === normalizeFieldName(names[nameIndex]))[0];
    if (exactKey) return row[exactKey];
  }
  for (let nameIndex = 0; nameIndex < names.length; nameIndex += 1) {
    const wanted = normalizeFieldName(names[nameIndex]);
    const field = fields.filter(candidate => normalizeFieldName(candidate.Title) === wanted || normalizeFieldName(candidate.InternalName) === wanted)[0];
    if (field && Object.prototype.hasOwnProperty.call(row, field.InternalName)) return row[field.InternalName];
  }
  const wanted = names.map(normalizeFieldName);
  const rowKeys = Object.keys(row);
  for (let index = 0; index < rowKeys.length; index += 1) {
    if (wanted.indexOf(normalizeFieldName(rowKeys[index])) >= 0) return row[rowKeys[index]];
  }
  return undefined;
};
const categoryNameFields: Record<string, string[]> = {
  'network-devices': ['DeviceName', 'Device Name', 'ProductFullName', 'Title', 'Name'],
  'asset-tracker': ['AssetDescription', 'Asset Description', 'ProductFullName', 'Title', 'Name'],
  'outsourced-services': ['Product/ServiceName', 'ProductServiceName', 'Product / Service Name', 'ServiceName', 'ProductFullName', 'Title', 'Name'],
  software: ['SoftwareFullName', 'Software Full Name', 'ProductFullName', 'Title', 'Name'],
  'hardware-server': ['ServerName', 'Server Name', 'ProductFullName', 'Title', 'Name'],
  'cloud-provider': ['ServiceName', 'Service Name', 'ProductFullName', 'Title', 'Name'],
  'hardware-accessories': ['ProductName', 'Product Name', 'ProductFullName', 'Title', 'Name'],
  'intellectual-property': ['AssetName', 'Asset Name', 'ProductFullName', 'Title', 'Name'],
  'media-devices': ['ProductFullName', 'Product Full Name', 'Title', 'Name'],
  'mobile-devices': ['ProductName', 'Product Name', 'ProductFullName', 'Title', 'Name'],
  'client-accessor': ['ProductName', 'Product Name', 'ProductFullName', 'Title', 'Name']
};
const displayRowValue = (value: unknown): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    const personValue = value as { Title?: string; LookupValue?: string; displayName?: string };
    return personValue.Title || personValue.LookupValue || personValue.displayName || '';
  }
  return String(value);
};
const mapSharePointAsset = (row: Record<string, unknown>, category: Category, fields: ISharePointField[], userNames: SharePointUserNameMap): Asset => {
  const valueFor = (names: string[]): unknown => getRowValue(row, names, fields);
  const textFor = (names: string[]): string => {
    const value = valueFor(names);
    if (typeof value === 'number' && names.some(name => /assignedto|owner|custodian/i.test(name))) return userNames[String(value)] || String(value);
    return displayRowValue(value);
  };
  const rawAmount = valueFor(['AssetValue', 'Value']);
  const amountRecord = rawAmount && typeof rawAmount === 'object' ? rawAmount as { Amount?: number; Value?: number } : undefined;
  const parsedAmount = Number(amountRecord?.Amount ?? amountRecord?.Value ?? rawAmount ?? 0);
  const nameFields = categoryNameFields[category.key] || ['ProductFullName', 'AssetDescription', 'Title', 'Name'];
  const mapped: Asset = {
    id: textFor(['AssetID', 'ID', 'Id']),
    name: textFor(nameFields) || '(Unnamed asset)',
    status: (textFor(['AssetStatus', 'Status']) || 'In Stock') as AssetStatus,
    assignedTo: textFor(['AssignedTo']) || textFor(['AssignedToId']).replace(/^\d+$/, id => userNames[id] || id) || 'Unassigned',
    purchaseDate: textFor(['DateOfPurchase', 'PurchaseDate']).slice(0, 10),
    value: Number.isFinite(parsedAmount) ? parsedAmount : 0,
    location: textFor(['AssetLocation', 'Location'])
  };
  mapped._spItemId = Number(valueFor(['ID', 'Id']));
  mapped._hasAttachments = valueFor(['Attachments']) === true || String(valueFor(['Attachments'])).toLowerCase() === 'true' ? 1 : 0;
  if (category.key === 'asset-tracker') {
    mapped.trackerAssetId = textFor(['AssetID', 'ID', 'Id']);
    mapped.trackerDescription = textFor(['AssetDescription']);
    mapped.trackerMovementDate = textFor(['DateOfMovement']);
    mapped.trackerMovedFrom = textFor(['AssetMovedFrom']);
    mapped.trackerMovedTo = textFor(['AssetMovedTo']);
    mapped.trackerStatus = textFor(['AssetStatus']);
    mapped.trackerAssignedTo = textFor(['AssignedTo']);
  }
  if (category.key === 'outsourced-services') {
    mapped.vendorId = textFor(['VendorID', 'Vendor ID']);
    mapped.vendorStatus = textFor(['VendorStatus', 'Vendor Status']);
    mapped.vendorFullName = textFor(['VendorFullName', 'Vendor Full Name', 'VendorName', 'Vendor Name']);
    mapped.productServiceName = textFor(['Product/ServiceName', 'ProductServiceName', 'Product / Service Name', 'ProductFullName', 'Title']);
    mapped.vendorPurpose = textFor(['Purpose']);
  }
  if (category.key === 'software') {
    mapped.softwareId = textFor(['SoftwareID', 'Software ID', 'ID', 'Id']);
    mapped.softwareFullName = textFor(['SoftwareFullName', 'Software Full Name', 'ProductFullName', 'Title']);
    mapped.softwareVersionModel = textFor(['Version/Model', 'VersionModel', 'Version', 'Model']);
    mapped.softwarePurposeService = textFor(['Purpose/Service', 'PurposeService', 'Purpose', 'Service Type']);
    mapped.softwarePlatform = textFor(['SoftwarePlatform', 'Software Platform']);
    mapped.softwareLicenseType = textFor(['Freeware/Licensed', 'FreewareLicensed', 'License Type', 'LicenseType']);
    mapped.softwareExpiry = textFor(['Licensed/SubscriptionExpiryDate', 'LicensedSubscriptionExpiryDate', 'License Expiry Date', 'Software Expiry Date', 'ExpiryDate']);
    mapped.softwareQtyPurchased = textFor(['Qty.Purchased', 'Qty Purchased', 'Quantity Purchased']);
    mapped.softwareQtyUsed = textFor(['Qty.Used', 'Qty Used', 'Quantity Used']);
  }
  if (category.key === 'hardware-server') {
    mapped.serverAssetId = textFor(['AssetID', 'Asset ID', 'ID', 'Id']);
    mapped.serverName = textFor(['ServerName', 'Server Name', 'ProductFullName', 'Title']);
    mapped.serverPurpose = textFor(['Purpose']);
    mapped.serverHostName = textFor(['HostName', 'Host Name']);
    mapped.serverSerialNumber = textFor(['SerialNo.', 'SerialNo', 'Serial No.', 'SerialNumber', 'Serial Number']);
    mapped.serverIpAddress = textFor(['IPAddress', 'IP Address', 'IP']);
    mapped.serverWarrantyStatus = textFor(['WarrantyStatus', 'Warranty Status']);
  }
  if (category.key === 'cloud-provider') {
    mapped.cloudVendorId = textFor(['VendorID', 'Vendor ID']);
    mapped.cloudServiceName = textFor(['ServiceName', 'Service Name', 'ProductFullName', 'Title']);
    mapped.cloudServicePartner = textFor(['ServicePartner', 'Service Partner', 'Cloud Service Partner']);
    mapped.cloudServiceType = textFor(['ServiceType', 'Service Type']);
    mapped.cloudPurpose = textFor(['Purpose']);
    mapped.cloudAgreementExpiry = textFor(['Contract/AgreementExpiryDate', 'ContractAgreementExpiryDate', 'Agreement Expiry Date', 'Contract End Date', 'ContractEnd']);
  }
  if (category.key === 'hardware-accessories') {
    mapped.accessoryAssetId = textFor(['AssetID', 'Asset ID', 'ID', 'Id']);
    mapped.accessoryProductName = textFor(['ProductName', 'Product Name', 'ProductFullName', 'Title']);
    mapped.accessoryCategory = textFor(['Category', 'AssetCategory']);
    mapped.accessoryAssignedTo = textFor(['AssignedTo', 'Assigned To']);
    mapped.accessorySerialNumber = textFor(['SerialNo.', 'SerialNo', 'Serial Number', 'SerialNumber']);
    mapped.accessoryModelNumber = textFor(['ModelNo.', 'ModelNo', 'Model Number', 'ModelNumber']);
    mapped.accessoryWarrantyStatus = textFor(['WarrantyStatus', 'Warranty Status']);
  }
  if (category.key === 'intellectual-property') {
    mapped.ipAssetId = textFor(['AssetID', 'Asset ID', 'ID', 'Id']);
    mapped.ipAssetName = textFor(['AssetName', 'Asset Name', 'ProductFullName', 'Title']);
    mapped.ipAssetDescription = textFor(['AssetDescription', 'Asset Description']);
    mapped.ipAssignedTo = textFor(['AssignedTo', 'Assigned To']);
    mapped.ipDateOfRenewal = textFor(['DateOfRenewal', 'Date Of Renewal', 'RenewalDate', 'Renewal Due']);
  }
  if (category.key === 'media-devices') {
    mapped.mediaAssetId = textFor(['AssetID', 'Asset ID', 'ID', 'Id']);
    mapped.mediaProductFullName = textFor(['ProductFullName', 'Product Full Name', 'Title']);
    mapped.mediaCapacity = textFor(['Capacity']);
    mapped.mediaSerialNumber = textFor(['SerialNo.', 'SerialNo', 'Serial Number', 'SerialNumber']);
    mapped.mediaModelNumber = textFor(['ModelNo.', 'ModelNo', 'Model Number', 'ModelNumber']);
    mapped.mediaType = textFor(['MediaType', 'Media Type']);
    mapped.mediaWarrantyStatus = textFor(['WarrantyStatus', 'Warranty Status']);
  }
  if (category.key === 'mobile-devices') {
    mapped.mobileAssetId = textFor(['AssetID', 'Asset ID', 'ID', 'Id']);
    mapped.mobileProductName = textFor(['ProductName', 'Product Name', 'ProductFullName', 'Title']);
    mapped.mobileAssignedTo = textFor(['AssignedTo', 'Assigned To']);
    mapped.mobileSerialNumber = textFor(['SerialNo.', 'SerialNo', 'Serial Number', 'SerialNumber', 'IMEI/Serial', 'IMEI']);
    mapped.mobileModelNumber = textFor(['ModelNo.', 'ModelNo', 'Model Number', 'ModelNumber']);
    mapped.mobileWarrantyStatus = textFor(['WarrantyStatus', 'Warranty Status']);
  }
  if (category.key === 'client-accessor') {
    mapped.clientAccessoryAssetId = textFor(['AssetID', 'Asset ID', 'ID', 'Id']);
    mapped.clientAccessoryProductName = textFor(['ProductName', 'Product Name', 'ProductFullName', 'Title']);
    mapped.clientAccessoryAssignedTo = textFor(['AssignedTo', 'Assigned To']);
    mapped.clientAccessorySerialNumber = textFor(['SerialNo.', 'SerialNo', 'Serial Number', 'SerialNumber']);
    mapped.clientAccessoryModelNumber = textFor(['ModelNo.', 'ModelNo', 'Model Number', 'ModelNumber']);
    mapped.clientAccessoryWarrantyStatus = textFor(['WarrantyStatus', 'Warranty Status']);
  }
  Object.keys(row).forEach(key => {
    const value = row[key];
    if (value && typeof value === 'object') {
      const record = value as { Email?: string; EMail?: string; email?: string; UserPrincipalName?: string; Id?: number; Title?: string; LookupValue?: string; Amount?: number; Value?: number; results?: Array<{ Email?: string; EMail?: string; email?: string; Title?: string }> };
      const field = fields.filter(candidate => normalizeFieldName(candidate.InternalName) === normalizeFieldName(key))[0];
      if (field?.TypeAsString === 'User') {
        mapped[key] = record.Email || record.EMail || record.email || record.UserPrincipalName || record.Title || '';
      } else if (field?.TypeAsString === 'UserMulti' && record.results) {
        mapped[key] = record.results.map(person => person.Email || person.EMail || person.email || person.Title || '').filter(Boolean).join(';#');
      } else {
        mapped[key] = record.Email || record.EMail || record.email || record.LookupValue || record.Title || record.Id || (record.Amount ?? record.Value ?? '');
      }
    } else if (value !== null && value !== undefined) {
      mapped[key] = typeof value === 'string' || typeof value === 'number' ? value : String(value);
    }
  });

  category.columns.forEach(column => {
    mapped[column.key] = displayRowValue(getRowValue(row, [column.key, column.label]));
  });
  if (category.key === 'network-devices') {
    mapped.networkSerialNumber = displayRowValue(getRowValue(row, ['SerialNumber', 'Serial No', 'Serial Number'], fields));
    mapped.networkModelNumber = displayRowValue(getRowValue(row, ['ModelNo', 'ModelNo.', 'Model Number'], fields));
    mapped.networkWarrantyStatus = displayRowValue(getRowValue(row, ['WarrantyStatus', 'Warranty Status'], fields));
  }
  category.fields.forEach(label => {
    mapped[fieldId(label)] = displayRowValue(getRowValue(row, [label]));
  });
  Object.keys(row).forEach(key => {
    if (key !== '__metadata' && key !== 'ID' && key !== 'Id') {
      const value = row[key];
      if (value !== null && value !== undefined) mapped[key] = displayRowValue(value);
    }
  });
  return mapped;
};
const AssetManagement: React.FC<IAssetManagementProps> = ({ userDisplayName, isFullScreen, isDarkTheme, spHttpClient, webAbsoluteUrl, graphClient }) => {
  const [categories, setCategories] = useState<Category[]>(emptyCategories);
  const [selectedKey, setSelectedKey] = useState('dashboard');
  const [dashboardItemIdsFilter, setDashboardItemIdsFilter] = useState<number[] | null>(null);
  const [dashboardCounts, setDashboardCounts] = useState<Record<string, number>>({});
  const [categoryQuery, setCategoryQuery] = useState('');
  const [query, setQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [page, setPage] = useState(1);
  const [modal, setModal] = useState<'view' | 'edit' | 'create' | null>(null);
  const [activeAssetId, setActiveAssetId] = useState<number | null>(null);
  const [viewAttachmentCount, setViewAttachmentCount] = useState(0);
  const [assetToDelete, setAssetToDelete] = useState<Asset | null>(null);
  const [draft, setDraft] = useState<Asset | null>(null);
  const [toast, setToast] = useState('');
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  const [isSaving, setIsSaving] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [createdAttachmentItemId, setCreatedAttachmentItemId] = useState<number | undefined>();
  const attachmentsRef = React.useRef<AssetAttachmentsHandle>(null);
  const [people, setPeople] = useState<IEntraEmployee[]>([]);
  const [peopleSearch, setPeopleSearch] = useState<Record<string, string>>({});
  const [debouncedPeopleSearch, setDebouncedPeopleSearch] = useState<Record<string, string>>({});
  const [peoplePickerOpen, setPeoplePickerOpen] = useState<Record<string, boolean>>({});
  const [isLoadingPeople, setIsLoadingPeople] = useState(false);
  const [peopleLoadError, setPeopleLoadError] = useState('');
  const [peopleReloadCount, setPeopleReloadCount] = useState(0);
  const [isLoadingAssets, setIsLoadingAssets] = useState(false);
  const [assetLoadError, setAssetLoadError] = useState('');
  const [sharePointFields, setSharePointFields] = useState<ISharePointField[]>([]);
  const category = categories.find(item => item.key === selectedKey) || categories[0];
  const orderedForm = orderAssetFormFields(selectedKey, sharePointFields);
  const networkDetailColumns = selectedKey === 'network-devices'
    ? [
      { key: 'networkSerialNumber', label: 'SerialNumber' },
      { key: 'networkModelNumber', label: 'ModelNo.' },
      { key: 'networkWarrantyStatus', label: 'WarrantyStatus' }
    ]
    : [];
  const laptopDetailColumns = selectedKey === 'hardware-laptop'
    ? ['Department', 'AssetSpecification', 'OperatingSystemDetails', 'WarrantyStatus'].map(label => ({ key: fieldId(label), label }))
    : [];
  const trackerDetailColumns = selectedKey === 'asset-tracker'
    ? [
      { key: 'trackerAssetId', label: 'AssetID' },
      { key: 'trackerDescription', label: 'AssetDescription' },
      { key: 'trackerMovementDate', label: 'DateOfMovement' },
      { key: 'trackerMovedFrom', label: 'AssetMovedFrom' },
      { key: 'trackerMovedTo', label: 'AssetMovedTo' },
      { key: 'trackerStatus', label: 'AssetStatus' },
      { key: 'trackerAssignedTo', label: 'AssignedTo' }
    ]
    : [];
  const outsourcedServiceColumns = selectedKey === 'outsourced-services'
    ? [
      { key: 'vendorId', label: 'VendorID' },
      { key: 'vendorStatus', label: 'VendorStatus' },
      { key: 'vendorFullName', label: 'VendorFullName' },
      { key: 'productServiceName', label: 'Product/ServiceName' },
      { key: 'vendorPurpose', label: 'Purpose' }
    ]
    : [];
  const softwareColumns = selectedKey === 'software'
    ? [
      { key: 'softwareId', label: 'SoftwareID' },
      { key: 'softwareFullName', label: 'SoftwareFullName' },
      { key: 'softwareVersionModel', label: 'Version/Model' },
      { key: 'softwarePurposeService', label: 'Purpose/Service' },
      { key: 'softwarePlatform', label: 'SoftwarePlatform' },
      { key: 'softwareLicenseType', label: 'Freeware/Licensed' },
      { key: 'softwareExpiry', label: 'Licensed/SubscriptionExpiryDate' },
      { key: 'softwareQtyPurchased', label: 'Qty.Purchased' },
      { key: 'softwareQtyUsed', label: 'Qty.Used' }
    ]
    : [];
  const serverColumns = selectedKey === 'hardware-server'
    ? [
      { key: 'serverAssetId', label: 'AssetID' },
      { key: 'serverName', label: 'ServerName' },
      { key: 'serverPurpose', label: 'Purpose' },
      { key: 'serverHostName', label: 'HostName' },
      { key: 'serverSerialNumber', label: 'SerialNo.' },
      { key: 'serverIpAddress', label: 'IPAddress' },
      { key: 'serverWarrantyStatus', label: 'WarrantyStatus' }
    ]
    : [];
  const cloudProviderColumns = selectedKey === 'cloud-provider'
    ? [
      { key: 'cloudVendorId', label: 'VendorID' },
      { key: 'cloudServiceName', label: 'ServiceName' },
      { key: 'cloudServicePartner', label: 'ServicePartner' },
      { key: 'cloudServiceType', label: 'ServiceType' },
      { key: 'cloudPurpose', label: 'Purpose' },
      { key: 'cloudAgreementExpiry', label: 'Contract/AgreementExpiryDate' }
    ]
    : [];
  const accessoryColumns = selectedKey === 'hardware-accessories'
    ? [
      { key: 'accessoryAssetId', label: 'AssetID' },
      { key: 'accessoryProductName', label: 'ProductName' },
      { key: 'accessoryCategory', label: 'Category' },
      { key: 'accessoryAssignedTo', label: 'AssignedTo' },
      { key: 'accessorySerialNumber', label: 'SerialNo.' },
      { key: 'accessoryModelNumber', label: 'ModelNo.' },
      { key: 'accessoryWarrantyStatus', label: 'WarrantyStatus' }
    ]
    : [];
  const intellectualPropertyColumns = selectedKey === 'intellectual-property'
    ? [
      { key: 'ipAssetId', label: 'AssetID' },
      { key: 'ipAssetName', label: 'AssetName' },
      { key: 'ipAssetDescription', label: 'AssetDescription' },
      { key: 'ipAssignedTo', label: 'AssignedTo' },
      { key: 'ipDateOfRenewal', label: 'DateOfRenewal' }
    ]
    : [];
  const mediaDeviceColumns = selectedKey === 'media-devices'
    ? [
      { key: 'mediaAssetId', label: 'AssetID' },
      { key: 'mediaProductFullName', label: 'ProductFullName' },
      { key: 'mediaCapacity', label: 'Capacity' },
      { key: 'mediaSerialNumber', label: 'SerialNo.' },
      { key: 'mediaModelNumber', label: 'ModelNo.' },
      { key: 'mediaType', label: 'MediaType' },
      { key: 'mediaWarrantyStatus', label: 'WarrantyStatus' }
    ]
    : [];
  const mobileDeviceColumns = selectedKey === 'mobile-devices'
    ? [
      { key: 'mobileAssetId', label: 'AssetID' },
      { key: 'mobileProductName', label: 'ProductName' },
      { key: 'mobileAssignedTo', label: 'Assigned To' },
      { key: 'mobileSerialNumber', label: 'SerialNo.' },
      { key: 'mobileModelNumber', label: 'ModelNo.' },
      { key: 'mobileWarrantyStatus', label: 'WarrantyStatus' }
    ]
    : [];
  const clientAccessoryColumns = selectedKey === 'client-accessor'
    ? [
      { key: 'clientAccessoryAssetId', label: 'AssetID' },
      { key: 'clientAccessoryProductName', label: 'ProductName' },
      { key: 'clientAccessoryAssignedTo', label: 'Assigned To' },
      { key: 'clientAccessorySerialNumber', label: 'SerialNo.' },
      { key: 'clientAccessoryModelNumber', label: 'ModelNo.' },
      { key: 'clientAccessoryWarrantyStatus', label: 'WarrantyStatus' }
    ]
    : [];
  const activeAsset = category.assets.find(asset => Number(asset._spItemId) === activeAssetId);
  const timelineEnabled = dashboardRegisters.find(item => item.key === selectedKey)?.timelineTab === true;
  const updateAttachmentFlag = (itemId: number, hasAttachments: boolean): void => {
    setCategories(current => current.map(item => item.key !== selectedKey ? item : {
      ...item,
      assets: item.assets.map(asset => Number(asset._spItemId) === itemId ? { ...asset, _hasAttachments: hasAttachments ? 1 : 0 } : asset)
    }));
  };
  const viewDetailsContent = activeAsset ? <dl className={styles.assetDetails}>{sharePointFields.length
    ? sharePointFields.map(field => <div key={field.InternalName}><dt>{field.Title}</dt><dd>{getAssetFieldValue(activeAsset, field)}</dd></div>)
    : <><div><dt>Asset ID</dt><dd className={styles.mono}>{activeAsset.id}</dd></div><div><dt>Status</dt><dd><span className={`${styles.status} ${getStatusClass(activeAsset.status)}`}>{activeAsset.status}</span></dd></div><div><dt>Assigned to</dt><dd>{activeAsset.assignedTo}</dd></div><div><dt>Location</dt><dd>{activeAsset.location}</dd></div><div><dt>Purchased</dt><dd>{activeAsset.purchaseDate}</dd></div><div><dt>Asset Value</dt><dd>{Number(activeAsset.value) || 0}</dd></div>{category.columns.map(column => <div key={column.key}><dt>{humanizeLabel(column.label)}</dt><dd>{activeAsset[column.key] || 'N/A'}</dd></div>)}</>}
  </dl> : null;
  const viewAttachmentsContent = activeAsset ? <AssetAttachments key={`${selectedKey}-${activeAsset._spItemId}-view`} listUrl={assetRegisterUrls[selectedKey]} webUrl={webAbsoluteUrl} client={spHttpClient} itemId={Number(activeAsset._spItemId)} mode="view" onCountChange={setViewAttachmentCount} onAttachmentsChanged={hasFiles => updateAttachmentFlag(Number(activeAsset._spItemId), hasFiles)} /> : null;
  const purchaseDate = activeAsset ? String(getRowValue(activeAsset, ['DateOfPurchase', 'PurchaseDate'], sharePointFields) || '') : '';
  const supportExpiryFields = dashboardRegisters.find(item => item.key === 'hardware-laptop')?.expiry?.[0]?.fields || [];
  const supportExpiryDate = activeAsset ? String(getRowValue(activeAsset, supportExpiryFields, sharePointFields) || '') : '';
  const visibleCategories = categories.filter(item => item.name.toLowerCase().includes(categoryQuery.toLowerCase()));
  const filteredAssets = useMemo(() => category.assets.filter(asset => {
    const normalizedQuery = query.trim().toLocaleLowerCase();
    const serialValues = Object.keys(asset).filter(key => /serial|imei/i.test(key)).map(key => String(asset[key] ?? '').toLocaleLowerCase());
    const matchesSearch = !normalizedQuery || Object.keys(asset).some(key => String(asset[key] ?? '').toLocaleLowerCase().includes(normalizedQuery)) || serialValues.some(value => value.includes(normalizedQuery));
    return (!dashboardItemIdsFilter || dashboardItemIdsFilter.indexOf(Number(asset._spItemId)) >= 0) &&
    matchesSearch &&
    (!statusFilter || getAssetFilterStatuses(asset, category.key).some(status => status.toLocaleLowerCase() === statusFilter.toLocaleLowerCase()));
  }), [category.assets, category.key, query, statusFilter, dashboardItemIdsFilter]);
  const pageSize = 6;
  const pageCount = Math.max(1, Math.ceil(filteredAssets.length / pageSize));
  const pageAssets = filteredAssets.slice((page - 1) * pageSize, page * pageSize);
  const categoryStatusCounts = category.assets.reduce<Record<string, number>>((counts, asset) => {
    const status = getCategoryPrimaryStatus(asset, category.key);
    if (status) {
      const existingStatus = Object.keys(counts).find(value => value.toLocaleLowerCase() === status.toLocaleLowerCase());
      const countKey = existingStatus || status;
      counts[countKey] = (counts[countKey] || 0) + 1;
    }
    return counts;
  }, {});
  const categoryStatusOptions = category.assets.reduce<string[]>((options, asset) => {
    const status = getCategoryPrimaryStatus(asset, category.key);
    if (status && !options.some(existing => existing.toLocaleLowerCase() === status.toLocaleLowerCase())) options.push(status);
    return options;
  }, []).concat(category.assets.reduce<string[]>((options, asset) => {
    getWarrantyStatusValues(asset).forEach(status => {
      if (!options.some(existing => existing.toLocaleLowerCase() === status.toLocaleLowerCase())) options.push(status);
    });
    return options;
  }, []));
  const categoryStatusMetrics = Object.keys(categoryStatusCounts)
    .map(status => ({ status, count: categoryStatusCounts[status] }))
    .sort((left, right) => right.count - left.count)
    .slice(0, 2);

  useEffect(() => {
    let active = true;
    if (selectedKey === 'dashboard') { setIsLoadingAssets(false); setAssetLoadError(''); return () => { active = false; }; }
    const item = initialCategories.find(candidate => candidate.key === selectedKey);
    if (!item) return () => { active = false; };
    setIsLoadingAssets(true);
    setAssetLoadError('');
    const listUrl = assetRegisterUrls[item.key];
    (async () => {
      const fields = await getSharePointListFields(spHttpClient, listUrl);
      const rows = await getSharePointListItems(spHttpClient, listUrl, fields);
      const needsUserNames = fields.some(field => field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti');
      const userNames = needsUserNames ? await getSharePointUserNames(spHttpClient, webAbsoluteUrl).catch(() => ({} as SharePointUserNameMap)) : {};
      if (!active) return;
      const assets = rows.map(row => mapSharePointAsset(row, item, fields, userNames));
      setCategories(current => current.map(currentCategory => currentCategory.key === item.key ? { ...currentCategory, assets } : currentCategory));
      setSharePointFields(fields); setAssetLoadError(''); setIsLoadingAssets(false);
    })().catch(error => { if (active) { setAssetLoadError(error instanceof Error ? error.message : `Unable to load ${item.name}.`); setIsLoadingAssets(false); } });
    return () => { active = false; };
  }, [selectedKey, spHttpClient, webAbsoluteUrl]);

  useEffect(() => {
    let active = true;
    setIsLoadingPeople(true);
    setPeopleLoadError('');
    getEntraEmployees(graphClient)
      .then(users => {
        if (active) {
          setPeople(users);
          setIsLoadingPeople(false);
        }
      })
      .catch(error => {
        if (active) {
          console.error('People directory request failed', error);
          setPeopleLoadError(error instanceof Error ? error.message : 'Unable to load employees.');
          setIsLoadingPeople(false);
        }
      });
    return () => { active = false; };
  }, [graphClient, peopleReloadCount]);

  useEffect(() => {
    const timeoutId = window.setTimeout(() => setDebouncedPeopleSearch(peopleSearch), 300);
    return () => window.clearTimeout(timeoutId);
  }, [peopleSearch]);

  const notify = (message: string): void => {
    setToast(message);
    window.setTimeout(() => setToast(''), 2400);
  };
  const selectCategory = (key: string): void => {
    setSelectedKey(key);
    setDashboardItemIdsFilter(null);
    setQuery('');
    setStatusFilter('');
    setPage(1);
    setModal(null);
  };
  const navigateFromDashboard = (key: string, filterStatus?: string, itemIds?: number[]): void => {
    setSelectedKey(key); setDashboardItemIdsFilter(itemIds || null); setQuery(''); setStatusFilter(filterStatus || ''); setPage(1); setModal(null);
  };
  const startCreate = (): void => {
    setCreatedAttachmentItemId(undefined);
    setFieldErrors({});
    const nextId = '';
    setActiveAssetId(null);
    const statusField = sharePointFields.find(field => normalizeFieldName(field.InternalName) === 'assetstatus' || normalizeFieldName(field.Title) === 'assetstatus');
    const configuredStatuses = statusField ? (Array.isArray(statusField.Choices) ? statusField.Choices : statusField.Choices?.results || []) : [];
    const defaultStatus = configuredStatuses[0] || categoryStatusOptions[0] || '';
    const newDraft: Asset = { id: nextId, name: '', status: defaultStatus as AssetStatus, assignedTo: '', purchaseDate: '', value: 0, location: '' };
    sharePointFields.forEach(field => {
      if (normalizeFieldName(field.InternalName) === 'assetid') newDraft[field.InternalName] = nextId;
      else if (normalizeFieldName(field.InternalName) === 'assetstatus' || normalizeFieldName(field.Title) === 'assetstatus') {
        newDraft[field.InternalName] = defaultStatus;
      }
      else if (normalizeFieldName(field.InternalName) === 'assetvalue') newDraft[field.InternalName] = ['Number', 'Currency', 'Integer'].indexOf(field.TypeAsString) >= 0 ? 0 : '';
      else if (field.TypeAsString !== 'User' && field.TypeAsString !== 'UserMulti') newDraft[field.InternalName] = '';
    });
    setDraft(newDraft);
    setModal('create');
  };
  const startEdit = (asset: Asset): void => {
    setActiveAssetId(Number(asset._spItemId));
    setFieldErrors({});
    setDraft({ ...asset });
    setModal('edit');
  };
  const updateDraft = (key: string, value: string): void => {
    setFieldErrors(current => {
      if (!current[key]) return current;
      const next = { ...current };
      delete next[key];
      return next;
    });
    setDraft(current => {
      if (!current) return current;
      return { ...current, [key]: key === 'value' ? Number(value) : value };
    });
  };
  const findPerson = (value: string): IEntraEmployee | undefined => {
    const normalized = value.trim().toLocaleLowerCase();
    return people.find(person => [person.mail, person.userPrincipalName, person.id, person.displayName]
      .some(identity => identity?.toLocaleLowerCase() === normalized));
  };
  const fieldErrorMessage = (key: string): React.ReactElement | null => fieldErrors[key]
    ? <small className={styles.formFieldError} id={`${key}-error`} role="alert">{fieldErrors[key]}</small>
    : null;
  const renderPeoplePicker = (key: string, label: string, nameKey: string, required = false, multiple = false): React.ReactElement => {
    const rawValue = String(draft?.[key] ?? '');
    const selectedValues = multiple ? rawValue.split(';#').filter(Boolean) : (rawValue ? [rawValue] : []);
    const selectedUsers = selectedValues.map(value => ({ value, person: findPerson(value) }));
    const selectedPerson = selectedUsers[0]?.person;
    const isOpen = !!peoplePickerOpen[key];
    const searchText = peopleSearch[key] || '';
    const debouncedSearchText = debouncedPeopleSearch[key] || '';
    const normalizedSearch = debouncedSearchText.trim().toLocaleLowerCase();
    const isSearching = searchText !== debouncedSearchText;
    const selectedIdentities = new Set(selectedValues.map(value => value.toLocaleLowerCase()));
    const matches = people.filter(person => {
      const identity = person.mail || person.userPrincipalName || person.id;
      if (multiple && selectedIdentities.has(identity.toLocaleLowerCase())) return false;
      if (!normalizedSearch) return true;
      return [person.displayName, person.mail, person.userPrincipalName]
        .some(value => value?.toLocaleLowerCase().includes(normalizedSearch));
    });
    const closePicker = (): void => {
      setPeoplePickerOpen(current => ({ ...current, [key]: false }));
      setPeopleSearch(current => ({ ...current, [key]: '' }));
    };
    const choosePerson = (person: IEntraEmployee): void => {
      const identity = person.mail || person.userPrincipalName || person.id;
      if (multiple) {
        const next = selectedValues.concat(identity).filter((value, index, values) => values.findIndex(item => item.toLocaleLowerCase() === value.toLocaleLowerCase()) === index);
        updateDraft(key, next.join(';#'));
      } else {
        updateDraft(key, identity);
        updateDraft(nameKey, person.displayName);
        closePicker();
      }
      setPeopleSearch(current => ({ ...current, [key]: '' }));
    };
    const removePerson = (value: string): void => {
      const next = selectedValues.filter(item => item.toLocaleLowerCase() !== value.toLocaleLowerCase());
      updateDraft(key, next.join(';#'));
    };
    const hasError = !!fieldErrors[key];
    return <div className={styles.formField} key={key}>
      <span>{label}{required && <b aria-hidden="true"> *</b>}</span>
      <div className={styles.peoplePicker}>
        {multiple && selectedUsers.length > 0 && <div className={styles.peopleChips}>
          {selectedUsers.map(({ value, person }) => <span className={styles.peopleChip} key={value}>
            <span>{person?.displayName || value}</span>
            <button type="button" aria-label={`Remove ${person?.displayName || value}`} disabled={isSaving} onClick={() => removePerson(value)}>×</button>
          </span>)}
        </div>}
        <div className={`${styles.peopleInputRow} ${hasError ? styles.fieldInputError : ''}`}>
          <input
            type="text"
            role="combobox"
            aria-label={label}
            aria-expanded={isOpen}
            aria-haspopup="listbox"
            aria-autocomplete="list"
            aria-required={required}
            aria-invalid={hasError}
            aria-describedby={hasError ? `${key}-error` : undefined}
            aria-controls={`${key}-people-options`}
            autoComplete="off"
            placeholder={isLoadingPeople ? 'Loading employees...' : peopleLoadError ? 'Unable to load employees' : people.length ? 'Search by name or email' : 'No employees available'}
            value={isOpen ? searchText : (!multiple ? selectedPerson?.displayName || (rawValue && !findPerson(rawValue) ? rawValue : '') : '')}
            disabled={isSaving}
            onFocus={() => setPeoplePickerOpen(current => ({ ...current, [key]: true }))}
            onChange={event => {
              // Capture the DOM value before scheduling a functional state update.
              // React may evaluate the updater after the synthetic event has ended.
              const nextSearchText = event.currentTarget.value;
              setPeopleSearch(current => ({ ...current, [key]: nextSearchText }));
              setPeoplePickerOpen(current => ({ ...current, [key]: true }));
            }}
            onBlur={() => window.setTimeout(closePicker, 150)}
            onKeyDown={event => {
              if (event.key === 'Escape') closePicker();
              if (event.key === 'Enter' && isOpen && !isSearching && matches.length) {
                event.preventDefault();
                choosePerson(matches[0]);
              }
              if (multiple && event.key === 'Backspace' && !searchText && selectedValues.length) removePerson(selectedValues[selectedValues.length - 1]);
              if (!multiple && event.key === 'Backspace' && !searchText && rawValue) {
                updateDraft(key, '');
                updateDraft(nameKey, '');
              }
            }}
          />
          {!multiple && rawValue && <button type="button" className={styles.peopleClear} aria-label={`Clear ${label}`} disabled={isSaving} onMouseDown={event => event.preventDefault()} onClick={() => {
            updateDraft(key, '');
            updateDraft(nameKey, '');
            setPeopleSearch(current => ({ ...current, [key]: '' }));
            setPeoplePickerOpen(current => ({ ...current, [key]: true }));
          }}>×</button>}
          <button type="button" className={styles.peopleToggle} aria-label={`${isOpen ? 'Close' : 'Open'} ${label} options`} disabled={isSaving} onMouseDown={event => event.preventDefault()} onClick={() => {
            setPeoplePickerOpen(current => ({ ...current, [key]: !current[key] }));
            setPeopleSearch(current => ({ ...current, [key]: '' }));
          }}><Icon iconName="ChevronDown" /></button>
        </div>
        {isOpen && <div className={styles.peopleOptions} id={`${key}-people-options`} role="listbox" aria-label={`${label} options`}>
          {isLoadingPeople || isSearching ? <div className={styles.peopleEmpty}><span className={styles.peopleSpinner} aria-hidden="true" />{isLoadingPeople ? 'Loading employees...' : 'Searching employees...'}</div>
            : peopleLoadError ? <div className={styles.peopleEmpty}>{peopleLoadError}<button type="button" className={styles.peopleRetry} onMouseDown={event => event.preventDefault()} onClick={() => setPeopleReloadCount(count => count + 1)}>Try again</button></div>
              : matches.length ? matches.map(person => {
              const identity = person.mail || person.userPrincipalName || person.id;
              return <button type="button" role="option" aria-selected={selectedIdentities.has(identity.toLocaleLowerCase())} className={styles.peopleOption} key={person.id} onMouseDown={event => event.preventDefault()} onClick={() => choosePerson(person)}>
                <span className={styles.peopleAvatar}>{person.displayName.split(/\s+/).map(part => part[0]).slice(0, 2).join('').toUpperCase()}</span>
                <span className={styles.peopleOptionText}><strong>{person.displayName}</strong><small>{person.mail || person.userPrincipalName}</small></span>
                {multiple && selectedIdentities.has(identity.toLocaleLowerCase()) && <Icon iconName="Accept" />}
              </button>;
            }) : <div className={styles.peopleEmpty}>{normalizedSearch ? 'No matching employees' : 'No employees available'}</div>}
        </div>}
        {fieldErrorMessage(key)}
      </div>
    </div>;
  };
  const renderSharePointField = (field: ISharePointField): React.ReactElement | null => {
    const fieldKey = field.InternalName;
    const fieldName = normalizeFieldName(fieldKey);
    const fieldTitleName = normalizeFieldName(field.Title);
    const value = draft?.[fieldKey] ?? '';
    const hasError = !!fieldErrors[fieldKey];
    const errorClass = hasError ? styles.fieldInputError : '';
    const validationError = fieldErrorMessage(fieldKey);
    const isRequired = field.Required || modal === 'create';
    const ciaFieldLabel = ciaFields.find(label =>
      normalizeFieldName(label) === fieldName || normalizeFieldName(label) === fieldTitleName
    );
    if (ciaFieldLabel) {
      const fieldChoices = Array.isArray(field.Choices) ? field.Choices : field.Choices?.results || [];
      return <label className={styles.formField} key={fieldKey}>
        <span>{humanizeLabel(field.Title)}{isRequired && <b aria-hidden="true"> *</b>}</span>
        <select className={errorClass} aria-required={isRequired} aria-invalid={hasError} aria-describedby={hasError ? `${fieldKey}-error` : undefined} value={String(value)} onChange={event => updateDraft(fieldKey, event.currentTarget.value)}>
          <option value="">Select</option>
          {fieldChoices.map(choice => <option key={choice} value={choice}>{choice}</option>)}
        </select>
        {validationError}
      </label>;
    }
    if (fieldName === 'assetstatus' || fieldTitleName === 'assetstatus') {
      const fieldChoices = Array.isArray(field.Choices) ? field.Choices : field.Choices?.results || [];
      const statusChoices = fieldChoices;
      return <label className={styles.formField} key={fieldKey}>
        <span>{humanizeLabel(field.Title)}{isRequired && <b aria-hidden="true"> *</b>}</span>
        <select className={errorClass} aria-required={isRequired} aria-invalid={hasError} aria-describedby={hasError ? `${fieldKey}-error` : undefined} value={String(value)} onChange={event => updateDraft(fieldKey, event.currentTarget.value)}>
          <option value="">Select</option>
          {statusChoices.map(choice => <option key={choice} value={choice}>{choice}</option>)}
        </select>
        {validationError}
      </label>;
    }
    if (field.TypeAsString === 'User') {
      return renderPeoplePicker(fieldKey, humanizeLabel(field.Title), `${fieldKey}DisplayName`, isRequired);
    }
    if (field.TypeAsString === 'UserMulti') {
      return renderPeoplePicker(fieldKey, humanizeLabel(field.Title), `${fieldKey}DisplayName`, isRequired, true);
    }
    if (field.TypeAsString === 'Choice') {
      const choices = Array.isArray(field.Choices) ? field.Choices : field.Choices?.results || [];
      return <label className={styles.formField} key={fieldKey}><span>{humanizeLabel(field.Title)}{isRequired && <b aria-hidden="true"> *</b>}</span><select className={errorClass} aria-required={isRequired} aria-invalid={hasError} aria-describedby={hasError ? `${fieldKey}-error` : undefined} value={value} onChange={event => updateDraft(fieldKey, event.currentTarget.value)}><option value="">Select</option>{choices.map(choice => <option key={choice}>{choice}</option>)}</select>{validationError}</label>;
    }
    if (field.TypeAsString === 'MultiChoice') {
      const choices = Array.isArray(field.Choices) ? field.Choices : field.Choices?.results || [];
      const selectedChoices = String(value).split(';#').filter(Boolean);
      return <label className={styles.formField} key={fieldKey}><span>{humanizeLabel(field.Title)}{isRequired && <b aria-hidden="true"> *</b>}</span><select className={errorClass} multiple aria-required={isRequired} aria-invalid={hasError} aria-describedby={hasError ? `${fieldKey}-error` : undefined} value={selectedChoices} onChange={event => updateDraft(fieldKey, Array.from(event.currentTarget.selectedOptions).map(option => option.value).join(';#'))}>{choices.map(choice => <option key={choice}>{choice}</option>)}</select>{validationError}</label>;
    }
    const type = field.TypeAsString === 'DateTime' ? 'date'
      : field.TypeAsString === 'Number' || field.TypeAsString === 'Currency' || field.TypeAsString === 'Integer' ? 'number'
        : field.TypeAsString === 'Boolean' ? 'checkbox' : 'text';
    const displayValue = type === 'date' && value ? String(value).slice(0, 10) : value;
    return <label className={styles.formField} key={fieldKey}>
      <span>{humanizeLabel(field.Title)}{isRequired && <b aria-hidden="true"> *</b>}</span>
      <input className={errorClass} aria-required={isRequired} aria-invalid={hasError} aria-describedby={hasError ? `${fieldKey}-error` : undefined} type={type} min={type === 'number' ? 0 : undefined} step={field.TypeAsString === 'Currency' ? '0.01' : undefined} checked={type === 'checkbox' ? String(value).toLowerCase() === 'true' : undefined} value={type === 'checkbox' ? undefined : displayValue} onChange={event => updateDraft(fieldKey, type === 'checkbox' ? String(event.currentTarget.checked) : event.currentTarget.value)} />
      {validationError}
    </label>;
  };
  const saveAsset = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    if (!draft || isSaving) return;

    const validPersonIdentities: string[] = [];
    people.forEach(person => {
      [person.mail, person.userPrincipalName, person.id].forEach(identity => {
        if (identity) validPersonIdentities.push(identity);
      });
    });
    const validationErrors = validateAssetForm(sharePointFields, draft, modal === 'create'
      ? { requireAllFields: true, validPersonIdentities }
      : undefined);
    setFieldErrors(validationErrors);
    if (Object.keys(validationErrors).length) {
      const hasBlankField = Object.keys(validationErrors).some(key => validationErrors[key].endsWith(' is required.'));
      notify(hasBlankField ? 'Please fill in all fields.' : 'Please select valid values for the highlighted fields.');
      return;
    }

    const sharePointValues: Record<string, string | number> = {};
    sharePointFields.forEach(field => {
      let value = draft[field.InternalName];
      if (typeof value === 'string' && (field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti')) {
        const values = value.split(';#').filter(Boolean).map(personValue => {
          const person = people.filter(candidate =>
            candidate.displayName.toLowerCase() === personValue.toLowerCase() ||
            candidate.mail?.toLowerCase() === personValue.toLowerCase() ||
            candidate.userPrincipalName?.toLowerCase() === personValue.toLowerCase() ||
            candidate.id === personValue
          )[0];
          return person ? person.mail || person.userPrincipalName || person.id : personValue;
        });
        value = values.join(';#');
      }
      if (value !== undefined) sharePointValues[field.InternalName] = value;
    });
    setIsSaving(true);
    try {
      if (modal === 'create' && createdAttachmentItemId && attachmentsRef.current) {
        await saveSharePointListItem(spHttpClient, webAbsoluteUrl, assetRegisterUrls[selectedKey], sharePointFields, sharePointValues, createdAttachmentItemId);
        const failed = await attachmentsRef.current.uploadStaged(createdAttachmentItemId);
        if (failed) { notify(`Asset created, ${failed} file(s) failed to upload`); setIsSaving(false); return; }
        setIsSaving(false); setModal(null); setCreatedAttachmentItemId(undefined); notify('Asset created'); return;
      }
      const itemId = activeAssetId ? Number(draft._spItemId) : undefined;
      if (activeAssetId && (!Number.isInteger(itemId) || (itemId as number) <= 0)) {
        throw new Error('Cannot update this asset because its SharePoint item ID is missing. Reload the register and try again.');
      }
      const savedItemId = await saveSharePointListItem(spHttpClient, webAbsoluteUrl, assetRegisterUrls[selectedKey], sharePointFields, sharePointValues, itemId);
      if (!activeAssetId && attachmentsRef.current?.hasStaged() && savedItemId) setCreatedAttachmentItemId(savedItemId);
      const [refreshedRows, refreshedUsers] = await Promise.all([
        getSharePointListItems(spHttpClient, assetRegisterUrls[selectedKey], sharePointFields),
        sharePointFields.some(field => field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti')
          ? getSharePointUserNames(spHttpClient, webAbsoluteUrl)
          : Promise.resolve({} as SharePointUserNameMap)
      ]);
      const refreshedAssets = refreshedRows.map(row => mapSharePointAsset(row, category, sharePointFields, refreshedUsers));
      setCategories(current => current.map(item => item.key === selectedKey ? { ...item, assets: refreshedAssets } : item));
      if (attachmentsRef.current?.hasStaged()) {
        const createdId = activeAssetId ? itemId : savedItemId;
        if (!createdId) throw new Error('The asset was saved, but SharePoint did not return its item ID. Reload the register before attaching files.');
        if (!activeAssetId) setCreatedAttachmentItemId(createdId);
        const failed = await attachmentsRef.current.uploadStaged(createdId);
        if (failed) { notify(`${activeAssetId ? 'Changes saved' : 'Asset created'}, ${failed} file(s) failed to upload`); setIsSaving(false); return; }
      }
    } catch (error) {
      console.error('Asset form save failed', error);
      const detail = error instanceof Error ? error.message.toLocaleLowerCase() : '';
      const message = detail.includes('permission') || detail.includes('401') || detail.includes('403')
        ? 'You do not have permission to save this asset. Check your SharePoint access and try again.'
        : detail.includes('could not resolve')
          ? 'The selected person could not be matched to a SharePoint user. Choose a valid person and try again.'
          : 'Unable to save this asset. Check the field values and your SharePoint connection, then try again.';
      notify(message);
      setIsSaving(false);
      return;
    }
    setIsSaving(false);
    setPage(1);
    setModal(null);
    notify(activeAssetId ? 'Changes saved' : 'Asset created');
  };
  const removeAsset = (asset: Asset): void => {
    setAssetToDelete(asset);
  };
  const confirmDeleteAsset = async (): Promise<void> => {
    if (!assetToDelete) return;
    const asset = assetToDelete;
    const itemId = Number(asset._spItemId);
    if (!Number.isInteger(itemId) || itemId <= 0) {
      notify('Cannot delete this item because its SharePoint item ID is missing. Reload the list and try again.');
      setAssetToDelete(null);
      return;
    }
    setIsDeleting(true);
    try {
      await deleteSharePointListItem(spHttpClient, assetRegisterUrls[selectedKey], itemId);
      const [rows, fields] = await Promise.all([
        getSharePointListItems(spHttpClient, assetRegisterUrls[selectedKey], sharePointFields),
        getSharePointListFields(spHttpClient, assetRegisterUrls[selectedKey])
      ]);
      const userNames = fields.some(field => field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti')
        ? await getSharePointUserNames(spHttpClient, webAbsoluteUrl)
        : {};
      const refreshedAssets = rows.map(row => mapSharePointAsset(row, category, fields, userNames));
      setSharePointFields(fields);
      setCategories(current => current.map(item => item.key === selectedKey ? { ...item, assets: refreshedAssets } : item));
      setModal(null);
      setAssetToDelete(null);
      notify('Asset deleted');
    } catch (error) {
      notify(error instanceof Error ? error.message : 'Unable to delete the SharePoint list item.');
    } finally {
      setIsDeleting(false);
    }
  };
  const exportExcel = async (): Promise<void> => {
    try {
      const listUrl = assetRegisterUrls[selectedKey];
      const fields = await getSharePointListFields(spHttpClient, listUrl);
      const rows = await getSharePointListItems(spHttpClient, listUrl, fields);
      const needsUserNames = fields.some(field => field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti');
      const userNames = needsUserNames ? await getSharePointUserNames(spHttpClient, webAbsoluteUrl) : {};
      const liveAssets = rows.map(row => mapSharePointAsset(row, category, fields, userNames));
      const matchingAssets = liveAssets.filter(asset =>
        Object.keys(asset).some(key => String(asset[key]).toLowerCase().includes(query.toLowerCase())) &&
        (!statusFilter || getAssetFilterStatuses(asset, category.key).some(status => status.toLocaleLowerCase() === statusFilter.toLocaleLowerCase()))
      );
      const columns = [
        { key: '_spItemId', label: 'ID' },
        ...fields.map(field => ({ key: field.InternalName, label: field.Title }))
      ];
      const workbook = await createExcelWorkbook(columns, matchingAssets);
      const url = URL.createObjectURL(workbook);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${category.key}-assets.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      notify(`${matchingAssets.length} assets exported from SharePoint`);
    } catch (error) {
      notify(error instanceof Error ? error.message : 'Unable to export items from the SharePoint list.');
    }
  };
  const setFilter = (value: string): void => {
    setStatusFilter(value);
    setPage(1);
  };
  const displayName = userDisplayName || 'Asset Manager';
  const initials = displayName.split(/\s+/).map(part => part[0]).slice(0, 2).join('').toUpperCase();
  return (
    <div className={`${styles.app} ${isFullScreen ? styles.fullScreen : ''} ${isDarkTheme ? styles.darkTheme : ''}`}>
      <aside className={styles.sidebar}>
        <div className={styles.brand}><span className={styles.brandIcon}><Icon iconName="Package" /></span><span><strong>Asset Management</strong><small>Asset register</small></span></div>
        <label className={styles.categorySearch}><Icon iconName="Search" /><input value={categoryQuery} onChange={event => setCategoryQuery(event.currentTarget.value)} placeholder="Find a register" /></label>
        <nav className={styles.categoryNav} aria-label="Asset registers">
          <button className={`${styles.categoryButton} ${selectedKey === 'dashboard' ? styles.categorySelected : ''}`} onClick={() => selectCategory('dashboard')}><Icon iconName="ViewDashboard"/><span>Dashboard</span></button>
          {visibleCategories.map(item => <button key={item.key} className={`${styles.categoryButton} ${item.key === selectedKey ? styles.categorySelected : ''}`} onClick={() => selectCategory(item.key)}>
            <Icon iconName={item.icon} /><span>{item.name}</span><small>{dashboardCounts[item.key] ?? item.assets.length}</small>
          </button>)}
          {!visibleCategories.length && <p className={styles.emptyNav}>No registers match.</p>}
        </nav>
        <div className={styles.profile}><span className={styles.avatar}>{initials}</span><span><strong>{displayName}</strong><small>Asset Manager</small></span></div>
      </aside>

      <div className={styles.workspace}>
        {selectedKey !== 'dashboard' && <header className={styles.header}>
          <div><h1>{category.name}</h1><p>{category.description}</p></div>
          <div className={styles.headerActions}><button className={styles.secondaryButton} onClick={exportExcel}><Icon iconName="ExcelDocument" />Export Excel</button><button className={styles.primaryButton} onClick={startCreate}><Icon iconName="Add" />Add asset</button></div>
        </header>}
        <main className={styles.main}>
          {selectedKey === 'dashboard' ? <Dashboard client={spHttpClient} onNavigate={navigateFromDashboard} onRegisterCounts={setDashboardCounts}/> : <>
          <section className={styles.metrics} aria-label="Asset summary">
            <div><span>Total assets</span><strong>{category.assets.length}</strong></div>
            {categoryStatusMetrics.map(metric => <div key={metric.status}><span>{metric.status}</span><strong className={getMetricClass(metric.status)}>{metric.count}</strong></div>)}
          </section>
          <section className={styles.tablePanel}>
            <div className={styles.filters}>
              <label className={styles.assetSearch}><Icon iconName="Search" /><input value={query} onChange={event => { setQuery(event.currentTarget.value); setPage(1); }} placeholder="Search by name, ID, assignee or serial number" /></label>
              <select value={statusFilter} onChange={event => setFilter(event.currentTarget.value)} aria-label="Filter by status"><option value="">All statuses</option>{categoryStatusOptions.map(status => <option key={status} value={status}>{status}</option>)}</select>
              <button className={styles.resetButton} onClick={() => { setQuery(''); setFilter(''); }}><Icon iconName="Refresh" />Reset</button>
            </div>
            <div className={styles.tableScroll}><table><thead><tr>{selectedKey === 'asset-tracker' ? trackerDetailColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'outsourced-services' ? outsourcedServiceColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'software' ? softwareColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'hardware-server' ? serverColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'cloud-provider' ? cloudProviderColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'hardware-accessories' ? accessoryColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'intellectual-property' ? intellectualPropertyColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'media-devices' ? mediaDeviceColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'mobile-devices' ? mobileDeviceColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : selectedKey === 'client-accessor' ? clientAccessoryColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>) : <><th>Asset</th><th>Status</th>{selectedKey !== 'network-devices' && <th>Assigned to</th>}{selectedKey !== 'hardware-laptop' && selectedKey !== 'network-devices' && <><th>Location</th><th>Purchased</th><th className={styles.alignRight}>Value</th></>}{selectedKey !== 'network-devices' && category.columns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>)}{laptopDetailColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>)}{networkDetailColumns.map(column => <th key={column.key}>{humanizeLabel(column.label)}</th>)}</>}<th><span className={styles.srOnly}>Actions</span></th></tr></thead>
              <tbody>{pageAssets.map(asset => <tr key={String(asset._spItemId)} onDoubleClick={() => { setActiveAssetId(Number(asset._spItemId)); setModal('view'); }}>
                {selectedKey === 'asset-tracker' ? trackerDetailColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'outsourced-services' ? outsourcedServiceColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'software' ? softwareColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'hardware-server' ? serverColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'cloud-provider' ? cloudProviderColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'hardware-accessories' ? accessoryColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'intellectual-property' ? intellectualPropertyColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'media-devices' ? mediaDeviceColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'mobile-devices' ? mobileDeviceColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : selectedKey === 'client-accessor' ? clientAccessoryColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>) : <><td><div className={styles.assetCell}><span><strong>{asset.name}</strong><small>{asset.id}</small></span></div></td>
                <td><span className={`${styles.status} ${getStatusClass(asset.status)}`}>{asset.status}</span></td>
                {selectedKey !== 'network-devices' && <td>{asset.assignedTo}</td>}{selectedKey !== 'hardware-laptop' && selectedKey !== 'network-devices' && <><td>{asset.location}</td><td>{asset.purchaseDate}</td><td className={styles.alignRight}>{Number(asset.value) || 0}</td></>}
                {selectedKey !== 'network-devices' && category.columns.map(column => <td className={styles.mono} key={column.key}>{renderAssetColumnValue(asset, column)}</td>)}
                {laptopDetailColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>)}
                {networkDetailColumns.map(column => <td key={column.key}>{renderAssetColumnValue(asset, column)}</td>)}</>}
                <td><div className={styles.rowActions}><button aria-label={`View ${asset.id}`} title="View asset" onClick={() => { setActiveAssetId(Number(asset._spItemId)); setModal('view'); }}><Icon iconName="View" /></button><button aria-label={`Edit ${asset.id}`} title="Edit asset" onClick={() => startEdit(asset)}><Icon iconName="Edit" /></button><button aria-label={`Delete ${asset.id}`} title="Delete asset" onClick={async () => { setActiveAssetId(Number(asset._spItemId)); await removeAsset(asset); }}><Icon iconName="Delete" /></button></div></td>
              </tr>)}
              {!pageAssets.length && <tr><td colSpan={selectedKey === 'asset-tracker' ? trackerDetailColumns.length + 1 : selectedKey === 'outsourced-services' ? outsourcedServiceColumns.length + 1 : selectedKey === 'software' ? softwareColumns.length + 1 : selectedKey === 'hardware-server' ? serverColumns.length + 1 : selectedKey === 'cloud-provider' ? cloudProviderColumns.length + 1 : selectedKey === 'hardware-accessories' ? accessoryColumns.length + 1 : selectedKey === 'intellectual-property' ? intellectualPropertyColumns.length + 1 : selectedKey === 'media-devices' ? mediaDeviceColumns.length + 1 : selectedKey === 'mobile-devices' ? mobileDeviceColumns.length + 1 : selectedKey === 'client-accessor' ? clientAccessoryColumns.length + 1 : selectedKey === 'network-devices' ? 6 : (selectedKey === 'hardware-laptop' ? 4 : 7) + category.columns.length + laptopDetailColumns.length} className={styles.emptyTable}><Icon iconName="FabricFolder" /><strong>{isLoadingAssets ? 'Loading assets...' : assetLoadError ? 'Could not load this register' : 'No assets found'}</strong><span>{assetLoadError || (isLoadingAssets ? 'Fetching items from SharePoint.' : 'This SharePoint list has no items matching your filters.')}</span></td></tr>}</tbody>
            </table></div>
            <div className={styles.pagination}><span>Showing {filteredAssets.length ? (page - 1) * pageSize + 1 : 0}-{Math.min(page * pageSize, filteredAssets.length)} of {filteredAssets.length}</span><div><button disabled={page <= 1} onClick={() => setPage(current => current - 1)}>Previous</button><span>{page} / {pageCount}</span><button disabled={page >= pageCount} onClick={() => setPage(current => current + 1)}>Next</button></div></div>
          </section>
          </>}
        </main>
      </div>

      {modal && <div className={styles.modalBackdrop} role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) setModal(null); }}>
        <section className={styles.modal} role="dialog" aria-modal="true" aria-labelledby="modal-title">
          <header className={styles.modalHeader}><h2 id="modal-title">{modal === 'view' ? activeAsset?.name : modal === 'edit' ? `Edit ${activeAsset?.name}` : `Add asset · ${category.name}`}</h2><button aria-label="Close dialog" onClick={() => setModal(null)}><Icon iconName="Cancel" /></button></header>
          {modal === 'view' && activeAsset && (timelineEnabled
            ? <LaptopViewTabs itemId={Number(activeAsset._spItemId)} purchaseDate={purchaseDate} expiryDate={supportExpiryDate} client={spHttpClient} details={viewDetailsContent} attachments={viewAttachmentsContent} attachmentCount={viewAttachmentCount}/>
            : <div className={styles.modalBody}>
              {viewDetailsContent}
              {viewAttachmentsContent}
              <div className={styles.history}>
                <div className={styles.historyHeading}><h3>Service history</h3><span>0 requests</span></div>
                <p className={styles.timelineEmpty}>No service history records found.</p>
              </div>
            </div>)}
          {(modal === 'create' || modal === 'edit') && draft && <form className={styles.modalBody} onSubmit={saveAsset}>
            {!!orderedForm.missing.length && <p className={styles.formFieldWarning}>Fields not found in this SharePoint register: {orderedForm.missing.join(', ')}.</p>}
            <div className={styles.formGrid}>
              {orderedForm.fields.map(renderSharePointField)}
            </div>
            <AssetAttachments ref={attachmentsRef} key={`${selectedKey}-${modal}`} listUrl={assetRegisterUrls[selectedKey]} webUrl={webAbsoluteUrl} client={spHttpClient} itemId={modal === 'edit' ? Number(draft._spItemId) : createdAttachmentItemId} mode={modal === 'create' ? createdAttachmentItemId ? 'edit' : 'add' : 'edit'} onAttachmentsChanged={hasFiles => { if (modal === 'edit') updateAttachmentFlag(Number(draft._spItemId), hasFiles); }} />
            <footer className={styles.formFooter}><button type="button" className={styles.secondaryButton} disabled={isSaving} onClick={() => { setModal(null); setCreatedAttachmentItemId(undefined); }}>Cancel</button><button type="submit" className={styles.primaryButton} disabled={isSaving}>{isSaving ? createdAttachmentItemId ? 'Retrying uploads...' : 'Saving...' : modal === 'edit' ? 'Save changes' : createdAttachmentItemId ? 'Retry failed uploads' : 'Create asset'}</button></footer>
          </form>}
          {modal === 'view' && activeAsset && <footer className={styles.viewFooter}><button className={styles.secondaryButton} onClick={() => setModal(null)}>Close</button><button className={styles.secondaryButton} onClick={() => removeAsset(activeAsset)}><Icon iconName="Delete" />Delete asset</button><button className={styles.primaryButton} onClick={() => startEdit(activeAsset)}><Icon iconName="Edit" />Edit asset</button></footer>}
        </section>
      </div>}
      {assetToDelete && <div className={styles.confirmBackdrop} role="presentation" onMouseDown={event => { if (!isDeleting && event.target === event.currentTarget) setAssetToDelete(null); }}>
        <section className={styles.confirmDialog} role="alertdialog" aria-modal="true" aria-labelledby="delete-dialog-title" aria-describedby="delete-dialog-description">
          <div className={styles.confirmIcon}><Icon iconName="Delete" /></div>
          <h2 id="delete-dialog-title">Delete asset?</h2>
          <p id="delete-dialog-description">Delete <strong>{assetToDelete.name}</strong> from the {category.name} SharePoint list? Attached files will be permanently deleted with this record. SharePoint recycle bin may allow recovery.</p>
          <footer className={styles.confirmActions}>
            <button type="button" className={styles.secondaryButton} disabled={isDeleting} onClick={() => setAssetToDelete(null)}>Cancel</button>
            <button type="button" className={styles.deleteButton} disabled={isDeleting} onClick={confirmDeleteAsset}>{isDeleting ? 'Deleting...' : 'Delete asset'}</button>
          </footer>
        </section>
      </div>}
      {toast && <div className={styles.toast} role="status">{toast}</div>}
    </div>
  );
};

export default AssetManagement;


