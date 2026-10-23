import { assetFormFieldAliases, orderAssetFormFields } from './assetFormFieldOrder';
import type { ISharePointField } from './hardwareLaptopSharePointService';

const field = (name: string): ISharePointField => ({
  Title: name,
  InternalName: name.replace(/[^a-z0-9]/gi, ''),
  TypeAsString: 'Text',
  Hidden: false,
  ReadOnlyField: false,
  Required: false
});

describe('asset form field ordering', () => {
  it('orders the configured fields for every register and retains unspecified available fields afterward', () => {
    Object.keys(assetFormFieldAliases).forEach(registerKey => {
      const requested = assetFormFieldAliases[registerKey].map(aliases => field(aliases[0]));
      const extra = field('Additional register field');
      const ordered = orderAssetFormFields(registerKey, [extra, ...requested]);
      expect(ordered.fields.map(item => item.Title)).toEqual(requested.map(item => item.Title).concat(extra.Title));
      expect(ordered.missing).toEqual([]);
    });
  });

  it('reports requested fields that are absent without creating them', () => {
    const result = orderAssetFormFields('hardware-laptop', [field('Asset Status'), field('AssetID')]);
    expect(result.fields.map(item => item.Title)).toEqual(['AssetID', 'Asset Status']);
    expect(result.missing).toContain('Product Full Name');
  });

  it('matches the affected SharePoint internal names to their display-order slots', () => {
    const cases: Array<[string, string[]]> = [
      ['cloud-provider', ['CloudPlatform', 'Purpose']],
      ['hardware-accessories', ['SerialNumber']],
      ['media-devices', ['SerialNumber']],
      ['mobile-devices', ['AssetOwner', 'AssetCustodian', 'MaximumOS']],
      ['client-accessor', ['SerialNumber', 'MaximumOS']]
    ];
    cases.forEach(([registerKey, names]) => {
      const resolved = orderAssetFormFields(registerKey, names.map(field));
      expect(resolved.fields.map(item => item.InternalName)).toEqual(names.map(name => field(name).InternalName));
      names.forEach(name => expect(resolved.missing).not.toContain(name));
    });
  });

  it('uses CloudPlatform in the first Cloud Providers form slot without requesting Cloud Service Provider', () => {
    const result = orderAssetFormFields('cloud-provider', [field('CloudPlatform'), field('Purpose')]);

    expect(result.fields[0].InternalName).toBe('CloudPlatform');
    expect(result.missing).not.toContain('Cloud Service Provider');
  });
});
