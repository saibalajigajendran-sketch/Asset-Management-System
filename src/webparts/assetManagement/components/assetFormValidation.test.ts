import { validateAssetForm } from './assetFormValidation';
import type { ISharePointField } from './hardwareLaptopSharePointService';

const field = (internalName: string, type: string, required: boolean, choices?: string[]): ISharePointField => ({
  Title: internalName,
  InternalName: internalName,
  TypeAsString: type,
  Hidden: false,
  ReadOnlyField: false,
  Required: required,
  Choices: choices
});

describe('asset form validation', () => {
  it('validates only required fields configured by the current register', () => {
    const fields = [field('AssetID', 'Text', true), field('Purpose', 'Text', false)];
    expect(validateAssetForm(fields, { AssetID: '', Purpose: '' })).toEqual({ AssetID: 'AssetID is required.' });
    expect(validateAssetForm([field('Purpose', 'Text', false)], { Purpose: '' })).toEqual({});
  });

  it('requires every existing field on Add and only accepts selected People Picker identities', () => {
    const fields = [field('SoftwareID', 'Text', false), field('AssetOwner', 'User', false), field('AssetValue', 'Choice', false, ['1', '2'])];
    expect(validateAssetForm(fields, { SoftwareID: ' SW-1 ', AssetOwner: 'typed name', AssetValue: '1' }, {
      requireAllFields: true,
      validPersonIdentities: ['owner@example.com']
    })).toEqual({ AssetOwner: 'Select a valid person for AssetOwner.' });
    expect(validateAssetForm(fields, { SoftwareID: '   ', AssetOwner: '', AssetValue: '' }, {
      requireAllFields: true,
      validPersonIdentities: ['owner@example.com']
    })).toEqual({
      SoftwareID: 'SoftwareID is required.',
      AssetOwner: 'AssetOwner is required.',
      AssetValue: 'AssetValue is required.'
    });
    expect(validateAssetForm(fields, { SoftwareID: 'SW-1', AssetOwner: 'owner@example.com', AssetValue: '1' }, {
      requireAllFields: true,
      validPersonIdentities: ['owner@example.com']
    })).toEqual({});
  });

  it('validates Choice and MultiChoice values against actual SharePoint choices', () => {
    const fields = [field('AssetValue', 'Choice', true, ['1', '2', '3']), field('Tags', 'MultiChoice', false, ['A', 'B'])];
    expect(validateAssetForm(fields, { AssetValue: '0', Tags: 'A;#Other' })).toEqual({
      AssetValue: 'Select a valid AssetValue.',
      Tags: 'Select valid options for Tags.'
    });
    expect(validateAssetForm(fields, { AssetValue: '2', Tags: 'A;#B' })).toEqual({});
    expect(validateAssetForm([field('AssetValue', 'Choice', false)], { AssetValue: '2' }).AssetValue)
      .toBe('Valid options for AssetValue are unavailable. Refresh the register and try again.');
    expect(validateAssetForm([field('Tags', 'MultiChoice', true, ['A', 'B'])], { Tags: ';#' }).Tags)
      .toBe('Tags is required.');
  });

  it('validates invalid numeric and date values before submission', () => {
    expect(validateAssetForm([field('Cost', 'Currency', false), field('PurchaseDate', 'DateTime', false)], {
      Cost: 'not money', PurchaseDate: 'not a date'
    })).toEqual({ Cost: 'Enter a valid number for Cost.', PurchaseDate: 'Enter a valid date for PurchaseDate.' });
  });
});
