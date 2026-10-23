jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));

import { buildDashboardQuery, calculateDashboardSummary, calculateDaysLeft, countStatuses, expiringAssets, expiryPillTone, stockLevel, type IDashboardAsset, type IDashboardData, type IDashboardRegisterData } from './dashboardDataService';
import { dashboardRegisters } from './dashboardConfig';
import type { ISharePointField } from './hardwareLaptopSharePointService';
import { buildSharePointItemProjection } from './hardwareLaptopSharePointService';

const asset = (overrides: Partial<IDashboardAsset> = {}): IDashboardAsset => ({ registerKey: 'hardware-laptop', registerName: 'Hardware - Laptops, Desktops', itemId: 1, assetId: 'L-1', name: 'Laptop', status: '', assignedTo: '', custodian: '', location: '', warrantyStatus: '', expiry: '', expiryType: 'Warranty', ...overrides });
const register = (key: string, assets: IDashboardAsset[], flags: Partial<IDashboardRegisterData> = {}): IDashboardRegisterData => ({ config: dashboardRegisters.find(item => item.key === key)!, assets, fields: [], hasStatus: false, hasQuantity: false, hasExpiry: false, ...flags });
const data = (...registers: IDashboardRegisterData[]): IDashboardData => ({ registers, updatedAt: new Date(2026, 0, 1) });

describe('dashboard calculations', () => {
  it('builds per-register selections only from fields declared for that register', () => {
    dashboardRegisters.forEach(config => {
      const aliases = [config.assetId, config.nameFields, config.status || [], config.warrantyStatus || [], ...(config.expiry || []).map(item => item.fields), config.assignedTo || [], config.custodian || [], config.location || [], config.totalCount || []].reduce<string[]>((all, group) => all.concat(group), []);
      const available: ISharePointField[] = aliases.map((name, index) => ({ Title: name, InternalName: `Field${index}`, TypeAsString: config.assignedTo?.indexOf(name) === 0 ? 'User' : 'Text', Hidden: false, ReadOnlyField: false, Required: false }));
      const query = buildDashboardQuery(config, available);
      expect(query.selectedFields.length).toBeGreaterThan(0);
      expect(query.selectedFields.every(field => available.some(item => item.InternalName === field))).toBe(true);
      expect(query.selectedFields.every(field => {
        const title = available.find(item => item.InternalName === field)!.Title.toLowerCase().replace(/[^a-z0-9]/g, '');
        return aliases.some(alias => alias.toLowerCase().replace(/[^a-z0-9]/g, '') === title);
      })).toBe(true);
      query.expand.forEach(personField => expect(query.select).toContain(`${personField}/Title`));
      query.expand.forEach(personField => expect(query.select).toContain(`${personField}/EMail`));
      const projection = buildSharePointItemProjection(available, query.selectedFields);
      expect(projection.select.split(',').every(part => part.trim().length > 0)).toBe(true);
      expect(projection.expand).toBe(query.expand.join(','));
    });
  });

  it('keeps each person projection as a separate select property and expands the field', () => {
    const custodian: ISharePointField = { Title: 'Asset Custodian', InternalName: 'AssetCustodian', TypeAsString: 'User', Hidden: false, ReadOnlyField: false, Required: false };
    const assignedTo: ISharePointField = { ...custodian, Title: 'Assigned To', InternalName: 'AssignedTo', TypeAsString: 'UserMulti' };
    const projection = buildSharePointItemProjection([custodian, assignedTo], ['AssetCustodian', 'AssignedTo']);
    expect(projection.select.split(',')).toEqual(['ID', 'Title', 'Attachments', 'AssetCustodian/Title', 'AssetCustodian/EMail', 'AssignedTo/Title', 'AssignedTo/EMail']);
    expect(projection.expand).toBe('AssetCustodian,AssignedTo');
  });

  it('calculates days at local midnight across day, month and year boundaries', () => {
    const today = new Date(2025, 11, 31, 15, 45);
    expect(calculateDaysLeft('2025-12-31', today)).toBe(0);
    expect(calculateDaysLeft('2026-01-01', today)).toBe(1);
    expect(calculateDaysLeft('2025-12-30', today)).toBe(-1);
    expect(calculateDaysLeft('2026-02-01', new Date(2026, 0, 31, 23, 50))).toBe(1);
    expect(calculateDaysLeft('2026-01-01T00:00:00Z', new Date(2026, 0, 1, 9))).toBe(0);
    expect(calculateDaysLeft('', today)).toBeNull();
  });

  it('uses the specified expiry pill thresholds', () => {
    expect(expiryPillTone(7)).toBe('red');
    expect(expiryPillTone(8)).toBe('amber');
    expect(expiryPillTone(15)).toBe('amber');
    expect(expiryPillTone(16)).toBe('blue');
    expect(expiryPillTone(30)).toBe('blue');
  });

  it('counts messy status strings case-insensitively and includes unspecified', () => {
    expect(countStatuses([asset({ status: ' allocated ' }), asset({ status: 'IN STOCK' }), asset({ status: ' In Maintenance ' }), asset({ status: 'retired' }), asset()])).toEqual({ Allocated: 1, 'In Stock': 1, 'In Maintenance': 1, Retired: 1, Unspecified: 1 });
  });

  it('filters expiry rows inclusively within each selected window', () => {
    const today = new Date(2026, 0, 1);
    const assets = [asset({ itemId: 1, expiry: '2026-01-01' }), asset({ itemId: 2, expiry: '2026-01-08' }), asset({ itemId: 3, expiry: '2026-01-16' }), asset({ itemId: 4, expiry: '2026-01-31' }), asset({ itemId: 5, expiry: '2025-12-31' }), asset({ itemId: 6 })];
    expect(expiringAssets(assets, 7, today).map(item => item.itemId)).toEqual([1, 2]);
    expect(expiringAssets(assets, 15, today).map(item => item.itemId)).toEqual([1, 2, 3]);
    expect(expiringAssets(assets, 30, today).map(item => item.itemId)).toEqual([1, 2, 3, 4]);
  });

  it('assigns stock levels using zero, 10 percent and configured minimum rules', () => {
    expect(stockLevel(0, 8, 2)).toBe('Very low');
    expect(stockLevel(1, 10, 2)).toBe('Very low');
    expect(stockLevel(2, 100, 2)).toBe('Very low');
    expect(stockLevel(3, 20, 2)).toBe('OK');
  });

  it('skips missing fields and surfaces only available status, expiry and stock data', () => {
    const missing = register('software', [asset({ registerKey: 'software', status: 'Allocated' })]);
    const available = register('hardware-laptop', [asset({ status: 'In Stock', expiry: '2026-01-10' })], { hasStatus: true, hasExpiry: true });
    const summary = calculateDashboardSummary(data(missing, available), 30, new Date(2026, 0, 1));
    expect(summary.total).toBe(2);
    expect(summary.statuses['In Stock']).toBe(1);
    expect(summary.statuses.Allocated).toBe(0);
    expect(summary.expiring).toHaveLength(1);
    expect(summary.stocks.map(item => item.config.key)).toEqual(['hardware-laptop']);
  });

  it('counts duplicate Asset IDs as separate records and reports the duplicate', () => {
    const first = asset({ itemId: 21, assetId: 'DUP-1' });
    const second = asset({ itemId: 22, assetId: 'DUP-1' });
    const laptops = register('hardware-laptop', [first], { hasStatus: true, hasExpiry: true });
    const network = register('network-devices', [second], { hasStatus: true });
    const summary = calculateDashboardSummary(data(laptops, network));
    expect(summary.total).toBe(2);
    expect(summary.duplicateIds).toEqual([{ id: 'DUP-1', count: 2 }]);
  });
});
