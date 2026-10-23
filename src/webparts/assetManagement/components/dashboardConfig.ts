export interface IDashboardRegisterConfig {
  key: string;
  name: string;
  listUrl: string;
  assetId: string[];
  nameFields: string[];
  status?: string[];
  warrantyStatus?: string[];
  expiry?: { fields: string[]; type: string }[];
  assignedTo?: string[];
  custodian?: string[];
  location?: string[];
  totalCount?: string[];
  minStock?: number;
  timelineTab?: boolean;
}

export const assetRegisterUrls: Record<string, string> = {
  'hardware-laptop': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_Hardware_Laptop/AllItems.aspx',
  'network-devices': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_NetworkDevices/AllItems.aspx',
  'asset-tracker': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_AssetTracker/AllItems.aspx',
  'outsourced-services': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_OutsourcedServices/AllItems.aspx',
  software: 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_Software/AllItems.aspx',
  'hardware-server': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_Hardware_Server/AllItems.aspx',
  'cloud-provider': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_Cloud_Providers/AllItems.aspx',
  'hardware-accessories': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_Hardware_Accessories/AllItems.aspx',
  'intellectual-property': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_Intellectual_Property/AllItems.aspx',
  'media-devices': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_MediaDevices/AllItems.aspx',
  'mobile-devices': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_Mobile_Devices/AllItems.aspx',
  'client-accessor': 'https://ittridentsqa.sharepoint.com/sites/TSQA-ITSupport/Lists/AssetRegister_Client_Accessories/AllItems.aspx'
};

const commonAssetId = ['AssetID', 'Asset ID'];
const supportExpiry = { fields: ['ExpiryDateofSupport', 'Expiry Date of Support', 'Support Expiry Date', 'SupportExpiryDate'], type: 'Warranty' };
const warrantyStatus = ['WarrantyStatus', 'Warranty Status'];

export const dashboardRegisters: IDashboardRegisterConfig[] = [
  { key: 'hardware-laptop', name: 'Hardware - Laptops, Desktops', listUrl: assetRegisterUrls['hardware-laptop'], assetId: commonAssetId, nameFields: ['ProductFullName', 'Product Name', 'Title'], status: ['AssetStatus', 'Asset Status'], warrantyStatus, expiry: [supportExpiry], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'], location: ['AssetLocation', 'Location'], minStock: 2, timelineTab: true },
  { key: 'network-devices', name: 'Network Devices', listUrl: assetRegisterUrls['network-devices'], assetId: commonAssetId, nameFields: ['DeviceName', 'ProductFullName', 'Title'], status: ['AssetStatus', 'Asset Status'], warrantyStatus, expiry: [supportExpiry], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'], location: ['AssetLocation', 'Location'], minStock: 2 },
  { key: 'asset-tracker', name: 'Asset tracker', listUrl: assetRegisterUrls['asset-tracker'], assetId: commonAssetId, nameFields: ['AssetDescription', 'ProductFullName', 'Title'], status: ['AssetStatus', 'Asset Status'], assignedTo: ['AssignedTo', 'Assigned To'], location: ['AssetMovedTo', 'Asset Moved To'], minStock: 2 },
  { key: 'outsourced-services', name: 'Outsourced Services', listUrl: assetRegisterUrls['outsourced-services'], assetId: commonAssetId, nameFields: ['Product/ServiceName', 'ProductServiceName', 'Title'], assignedTo: ['AssignedTo'], custodian: ['AssetCustodian'], location: ['Location'] },
  { key: 'software', name: 'Software', listUrl: assetRegisterUrls.software, assetId: ['SoftwareID', 'Software ID', ...commonAssetId], nameFields: ['SoftwareFullName', 'ProductFullName', 'Title'], expiry: [{ fields: ['License/Subscription Expiry Date', 'Licensed/SubscriptionExpiryDate', 'LicensedSubscriptionExpiryDate', 'License Expiry Date', 'ExpiryDate'], type: 'License' }], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'], location: ['AssetLocation', 'Location'] },
  { key: 'hardware-server', name: 'Hardware - Server', listUrl: assetRegisterUrls['hardware-server'], assetId: commonAssetId, nameFields: ['ServerName', 'ProductFullName', 'Title'], warrantyStatus, expiry: [supportExpiry], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'], location: ['AssetLocation', 'Location'], minStock: 2 },
  { key: 'cloud-provider', name: 'Cloud Providers', listUrl: assetRegisterUrls['cloud-provider'], assetId: commonAssetId, nameFields: ['ServiceName', 'ProductFullName', 'Title'], expiry: [{ fields: ['Contract/AgreementExpireDate', 'ContractAgreementExpireDate', 'Contract/AgreementExpiryDate', 'Agreement Expiry Date'], type: 'Contract' }], assignedTo: ['AssignedTo'], custodian: ['TechnicalOwner', 'Technical Owner'], location: ['Region', 'Location'] },
  { key: 'hardware-accessories', name: 'Hardware - Accessories', listUrl: assetRegisterUrls['hardware-accessories'], assetId: commonAssetId, nameFields: ['ProductName', 'ProductFullName', 'Title'], warrantyStatus, expiry: [supportExpiry], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'], location: ['AssetLocation', 'Location'], totalCount: ['TotalCount', 'Total Count'], minStock: 2 },
  { key: 'intellectual-property', name: 'Intellectual Property', listUrl: assetRegisterUrls['intellectual-property'], assetId: commonAssetId, nameFields: ['AssetName', 'ProductFullName', 'Title'], expiry: [{ fields: ['DateOfRenewal', 'Date Of Renewal', 'RenewalDate'], type: 'Renewal' }], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'] },
  { key: 'media-devices', name: 'Media Device', listUrl: assetRegisterUrls['media-devices'], assetId: commonAssetId, nameFields: ['ProductFullName', 'Title'], warrantyStatus, expiry: [supportExpiry], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'], location: ['AssetLocation', 'Location'], minStock: 2 },
  { key: 'mobile-devices', name: 'Mobile Devices', listUrl: assetRegisterUrls['mobile-devices'], assetId: commonAssetId, nameFields: ['ProductName', 'ProductFullName', 'Title'], warrantyStatus, expiry: [supportExpiry], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'], location: ['AssetLocation', 'Location'], minStock: 2 },
  { key: 'client-accessor', name: 'Client Accessories', listUrl: assetRegisterUrls['client-accessor'], assetId: commonAssetId, nameFields: ['ProductName', 'ProductFullName', 'Title'], warrantyStatus, expiry: [supportExpiry], assignedTo: ['AssignedTo', 'Assigned To'], custodian: ['AssetCustodian', 'Asset Custodian'], location: ['Location'], totalCount: ['TotalCount', 'Total Count'], minStock: 2 }
];

export const dashboardStockThresholds = dashboardRegisters.reduce<Record<string, number>>((result, register) => {
  if (register.minStock !== undefined) result[register.key] = register.minStock;
  return result;
}, {});
