jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));

import { shouldIncludeSharePointFormField, type ISharePointField } from './hardwareLaptopSharePointService';

const field = (internalName: string, title = internalName, flags: Partial<ISharePointField> = {}): ISharePointField => ({
  InternalName: internalName,
  Title: title,
  TypeAsString: 'Text',
  Hidden: true,
  ReadOnlyField: false,
  Required: false,
  ...flags
});

describe('SharePoint form field filtering', () => {
  it('retains requested hidden editable fields only on their configured registers', () => {
    const cases: Array<[string, string[]]> = [
      ['AssetRegister_Cloud_Providers', ['CloudPlatform', 'Purpose']],
      ['AssetRegister_Hardware_Accessories', ['SerialNumber']],
      ['AssetRegister_MediaDevices', ['SerialNumber']],
      ['AssetRegister_Mobile_Devices', ['AssetOwner', 'AssetCustodian', 'MaximumOS']],
      ['AssetRegister_Client_Accessories', ['SerialNumber', 'MaximumOS']]
    ];
    cases.forEach(([listName, fields]) => {
      fields.forEach(name => expect(shouldIncludeSharePointFormField(`https://tenant.example/sites/assets/Lists/${listName}/AllItems.aspx`, field(name))).toBe(true));
    });
    expect(shouldIncludeSharePointFormField('https://tenant.example/sites/assets/Lists/AssetRegister_Cloud_Providers/AllItems.aspx', field('CloudPlatform', 'CloudPlatform', { ReadOnlyField: true }))).toBe(false);
    expect(shouldIncludeSharePointFormField('https://tenant.example/sites/assets/Lists/AssetRegister_Cloud_Providers/AllItems.aspx', field('UnrelatedHiddenField'))).toBe(false);
  });
});
