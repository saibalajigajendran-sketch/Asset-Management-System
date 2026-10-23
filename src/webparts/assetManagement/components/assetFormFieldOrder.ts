import type { ISharePointField } from './hardwareLaptopSharePointService';

export interface IOrderedFormFields {
  fields: ISharePointField[];
  missing: string[];
}

export const assetFormFieldAliases: Record<string, string[][]> = {
  'hardware-laptop': [
    ['Asset ID', 'AssetID'], ['Asset Status', 'AssetStatus'], ['Product Full Name', 'ProductFullName'],
    ['Asset Owner', 'AssetOwner'], ['Asset Custodian', 'AssetCustodian'], ['Department'],
    ['Assigned To (User Name)', 'Assigned To', 'AssignedTo', 'User Name', 'UserName'], ['User ID', 'UserID'], ['User Role', 'UserRole'],
    ['Asset Location', 'AssetLocation'], ['Serial Number', 'SerialNumber'], ['Model No.', 'ModelNo'],
    ['Asset Specification', 'AssetSpecification'], ['Operating System Details', 'OperatingSystemDetails'],
    ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue'],
    ['Date Of Purchase', 'DateOfPurchase'], ['Warranty Status', 'WarrantyStatus'],
    ['Expiry Date of Support', 'ExpiryDateofSupport', 'Support Expiry Date', 'SupportExpiryDate']
  ],
  'hardware-accessories': [
    ['Asset ID', 'AssetID'], ['Product Name', 'ProductName'], ['Accessory Category', 'Category', 'AssetCategory'],
    ['Asset Owner', 'AssetOwner'], ['Asset Custodian', 'AssetCustodian'], ['Assigned To', 'AssignedTo'],
    ['Asset Location', 'AssetLocation'], ['Serial No', 'SerialNo', 'SerialNumber'], ['Model No.', 'ModelNo'],
    ['Total Count', 'TotalCount'], ['Warranty Status', 'WarrantyStatus'],
    ['Expiry Date of Support', 'ExpiryDateofSupport', 'Support Expiry Date', 'SupportExpiryDate'],
    ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'hardware-server': [
    ['Asset ID', 'AssetID'], ['Server Name', 'ServerName'], ['Asset Owner', 'AssetOwner'],
    ['Asset Custodian', 'AssetCustodian'], ['Purpose'], ['Host Name', 'HostName'], ['Asset Location', 'AssetLocation'],
    ['Serial Number', 'SerialNo', 'SerialNumber'], ['Model No.', 'ModelNo'], ['IP Address', 'IPAddress'],
    ['Warranty Status', 'WarrantyStatus'], ['Expiry Date of Support', 'ExpiryDateofSupport', 'Support Expiry Date', 'SupportExpiryDate'],
    ['Server Specifications', 'Server Specification', 'ServerSpecifications', 'Specifications', 'Processor / RAM'],
    ['Operating System Details', 'OperatingSystemDetails', 'Operating System'], ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'network-devices': [
    ['Asset ID', 'AssetID'], ['Asset Status', 'AssetStatus'], ['Product Full Name', 'ProductFullName', 'DeviceName'],
    ['Asset Owner', 'AssetOwner'], ['Asset Custodian', 'AssetCustodian'], ['Asset Location', 'AssetLocation'],
    ['Serial Number', 'SerialNo', 'SerialNumber'], ['Model No.', 'ModelNo'], ['Warranty Status', 'WarrantyStatus'],
    ['Expiry Date of Support', 'ExpiryDateofSupport', 'Support Expiry Date', 'SupportExpiryDate'],
    ['Specifications', 'Specification'], ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'mobile-devices': [
    ['Asset ID', 'AssetID'], ['Product Name', 'ProductName'], ['Asset Owner', 'AssetOwner'], ['Asset Custodian', 'AssetCustodian'],
    ['Assigned To', 'AssignedTo'], ['Asset Location', 'AssetLocation'], ['Serial Number', 'SerialNo', 'SerialNumber'],
    ['Model No.', 'ModelNo'], ['Current OS', 'CurrentOS'], ['MAX OS', 'MaximumOS', 'MAXOS', 'MaxOS'], ['Capacity'],
    ['Warranty Status', 'WarrantyStatus'], ['Expiry Date of Support', 'ExpiryDateofSupport', 'Support Expiry Date', 'SupportExpiryDate'],
    ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'client-accessor': [
    ['Asset ID', 'AssetID'], ['Product Name', 'ProductName'], ['Accessory Category', 'Category', 'AssetCategory'],
    ['Asset Owner', 'AssetOwner'], ['Asset Custodian', 'AssetCustodian'], ['Assigned To', 'AssignedTo'],
    ['Asset Location', 'AssetLocation', 'Location'], ['Serial No', 'SerialNo', 'SerialNumber'], ['Model No.', 'ModelNo'],
    ['Current OS', 'CurrentOS'], ['MAX OS', 'MaximumOS', 'MAXOS', 'MaxOS'], ['Total Count', 'TotalCount'], ['Warranty Status', 'WarrantyStatus'],
    ['Expiry Date of Support', 'ExpiryDateofSupport', 'Support Expiry Date', 'SupportExpiryDate'],
    ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  software: [
    ['Software ID', 'SoftwareID'], ['Software Full Name', 'SoftwareFullName', 'ProductFullName'], ['Version / Model', 'VersionModel', 'Version/Model'],
    ['Asset Owner', 'AssetOwner'], ['Asset Custodian', 'AssetCustodian'], ['Purpose / Service', 'PurposeService', 'Purpose/Service'],
    ['Software Platform', 'SoftwarePlatform'], ['Freeware / Licensed', 'FreewareLicensed', 'Freeware/Licensed'],
    ['License / Subscription Expiry Date', 'LicenseSubscriptionExpiryDate', 'LicensedSubscriptionExpiryDate'],
    ['Qty. Purchased', 'QtyPurchased'], ['Qty. Used', 'QtyUsed'], ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'media-devices': [
    ['Asset ID', 'AssetID'], ['Media Type', 'MediaType'], ['Product Full Name', 'ProductFullName'],
    ['Asset Owner', 'AssetOwner'], ['Asset Custodian', 'AssetCustodian'], ['Asset Location', 'AssetLocation'],
    ['Serial No', 'SerialNo', 'SerialNumber'], ['Model No.', 'ModelNo'], ['Capacity'], ['Warranty Status', 'WarrantyStatus'],
    ['Expiry Date of Support', 'ExpiryDateofSupport', 'Support Expiry Date', 'SupportExpiryDate'],
    ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'outsourced-services': [
    ['Vendor ID', 'VendorID'], ['Vendor Status', 'VendorStatus'], ['Vendor Full Name', 'VendorFullName'],
    ['Product / Service Name', 'ProductServiceName', 'Product/ServiceName'], ['Purpose'], ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'intellectual-property': [
    ['Asset ID', 'AssetID'], ['Asset Name', 'AssetName'], ['Asset Description', 'AssetDescription'], ['Asset Owner', 'AssetOwner'],
    ['Asset Custodian', 'AssetCustodian'], ['Assigned To (User Name)', 'Assigned To', 'AssignedTo', 'User Name', 'UserName'], ['Asset Location', 'AssetLocation', 'Location'],
    ['Date of Renewal', 'DateOfRenewal'], ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'cloud-provider': [
    ['Cloud Platform', 'CloudPlatform'], ['Cloud Service Partner', 'CloudServicePartner', 'ServicePartner'], ['Vendor ID', 'VendorID'],
    ['Service Name', 'ServiceName'], ['Service Type', 'ServiceType'], ['Purpose', 'Purpose / Use Case', 'PurposeUseCase', 'Purpose/UseCase'],
    ['Business Owner', 'BusinessOwner'], ['Technical Owner', 'TechnicalOwner'],
    ['Contract / Agreement Expiry Date', 'ContractAgreementExpiryDate', 'Contract/AgreementExpireDate'],
    ['Confidentiality'], ['Integrity'], ['Availability'], ['Asset Value', 'AssetValue']
  ],
  'asset-tracker': [
    ['Asset ID', 'AssetID'], ['Asset Description', 'AssetDescription'], ['Date of Movement', 'DateOfMovement'],
    ['Asset Moved From', 'AssetMovedFrom'], ['Asset Moved To', 'AssetMovedTo'], ['Asset Status', 'AssetStatus'],
    ['Assigned To', 'AssignedTo'], ['Moved By', 'MovedBy'], ['Verified By', 'VerifiedBy'], ['Approved By', 'ApprovedBy']
  ]
};

const normalize = (value: string): string => value.toLowerCase().replace(/[^a-z0-9]/g, '');

export const orderAssetFormFields = (registerKey: string, availableFields: ISharePointField[]): IOrderedFormFields => {
  const requestedOrder = assetFormFieldAliases[registerKey];
  if (!requestedOrder) return { fields: availableFields, missing: [] };
  const used = new Set<string>();
  const ordered: ISharePointField[] = [];
  const missing: string[] = [];
  requestedOrder.forEach(aliases => {
    const field = aliases.reduce<ISharePointField | undefined>((match, alias) => match || availableFields.find(candidate => !used.has(candidate.InternalName) &&
      (normalize(candidate.Title) === normalize(alias) || normalize(candidate.InternalName) === normalize(alias))), undefined);
    if (field) {
      used.add(field.InternalName);
      ordered.push(field);
    } else {
      missing.push(aliases[0]);
    }
  });
  return { fields: ordered.concat(availableFields.filter(field => !used.has(field.InternalName))), missing };
};
