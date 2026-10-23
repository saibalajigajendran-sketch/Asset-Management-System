import { SPHttpClient } from '@microsoft/sp-http';
import { dashboardRegisters, dashboardStockThresholds, type IDashboardRegisterConfig } from './dashboardConfig';
import { getSharePointListFields, getSharePointListItems, type ISharePointField } from './hardwareLaptopSharePointService';

export interface IDashboardAsset {
  registerKey: string; registerName: string; itemId: number; assetId: string; name: string; status: string;
  assignedTo: string; custodian: string; location: string; warrantyStatus: string; expiry: string; expiryType: string; totalCount?: number;
}
export interface IDashboardRegisterData { config: IDashboardRegisterConfig; assets: IDashboardAsset[]; fields: string[]; hasStatus: boolean; hasQuantity: boolean; hasExpiry: boolean; error?: string; warning?: string; }
export interface IDashboardData { registers: IDashboardRegisterData[]; updatedAt: Date; }
export type StockLevel = 'OK' | 'Low' | 'Very low';
export type ExpiryPillTone = 'red' | 'amber' | 'blue';

const normalize = (value: string): string => value.replace(/_x([0-9a-f]{4})_/gi, (_match, hex: string) => String.fromCharCode(parseInt(hex, 16))).toLowerCase().replace(/[^a-z0-9]/g, '');
const textValue = (value: unknown): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    const record = value as { Title?: string; LookupValue?: string; EMail?: string; Email?: string; displayName?: string; results?: Array<{ Title?: string; EMail?: string; Email?: string }>; value?: Array<{ Title?: string; EMail?: string; Email?: string }>; d?: { results?: Array<{ Title?: string; EMail?: string; Email?: string }> } };
    const people = record.results || record.value || record.d?.results;
    return record.Title || record.LookupValue || record.displayName || record.EMail || record.Email || people?.map(item => item.Title || item.EMail || item.Email || '').filter(Boolean).join('; ') || '';
  }
  return String(value).trim();
};
export const resolveDashboardFields = (config: IDashboardRegisterConfig, fields: ISharePointField[]): string[] => {
  const aliases = [config.assetId, config.nameFields, config.status || [], config.warrantyStatus || [], ...(config.expiry || []).map(item => item.fields), config.assignedTo || [], config.custodian || [], config.location || [], config.totalCount || []].reduce<string[]>((all, group) => all.concat(group), []);
  const selected: string[] = [];
  aliases.forEach(alias => {
    const key = normalize(alias);
    const field = fields.find(item => normalize(item.Title) === key || normalize(item.InternalName) === key);
    if (field && selected.indexOf(field.InternalName) < 0) selected.push(field.InternalName);
  });
  return selected;
};
export const buildDashboardQuery = (config: IDashboardRegisterConfig, fields: ISharePointField[]): { select: string[]; expand: string[]; selectedFields: string[] } => {
  const selectedFields = resolveDashboardFields(config, fields);
  const people = fields.filter(field => selectedFields.indexOf(field.InternalName) >= 0 && (field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti'));
  return {
    select: ['ID', 'Title', 'Attachments', ...selectedFields.filter(name => !people.some(field => field.InternalName === name)), ...people.reduce<string[]>((all, field) => all.concat([`${field.InternalName}/Title`, `${field.InternalName}/EMail`]), [])],
    expand: people.map(field => field.InternalName),
    selectedFields
  };
};
const metadataAndItems = async (client: SPHttpClient, config: IDashboardRegisterConfig): Promise<IDashboardRegisterData> => {
  const fields = await getSharePointListFields(client, config.listUrl);
  const fieldFor = (aliases: string[] = []): string | undefined => {
    for (const alias of aliases) {
      const target = normalize(alias);
      const field = fields.find(item => normalize(item.Title) === target || normalize(item.InternalName) === target);
      if (field) return field.InternalName;
    }
    return undefined;
  };
  let selectedFields = buildDashboardQuery(config, fields).selectedFields;
  let rows: Record<string, unknown>[];
  let warning: string | undefined;
  try { rows = await getSharePointListItems(client, config.listUrl, fields, selectedFields); }
  catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const missing = message.match(/(?:field|property).*?['"]([^'"]+)['"].*(?:does not exist|was not found|could not be found)/i);
    const badField = missing?.[1];
    const known = fields.map(field => field.InternalName);
    if (!badField && !/(?:field|property).*(?:does not exist|was not found|could not be found)/i.test(message)) throw error;
    const skipped = badField || 'unknown field';
    selectedFields = selectedFields.filter(name => known.indexOf(name) >= 0 && normalize(name) !== normalize(skipped));
    warning = `SharePoint skipped unavailable field "${skipped}" and retried this register.`;
    rows = await getSharePointListItems(client, config.listUrl, fields, selectedFields);
  }
  const get = (row: Record<string, unknown>, aliases: string[] = []): string => {
    const field = fieldFor(aliases);
    if (field && row[field] !== undefined) return textValue(row[field]);
    for (const alias of aliases) {
      const matchKey = Object.keys(row).find(key => normalize(key) === normalize(alias));
      if (matchKey) return textValue(row[matchKey]);
    }
    return '';
  };
  const assets: IDashboardAsset[] = rows.map(row => {
    const expiryConfig = (config.expiry || []).find(item => !!fieldFor(item.fields));
    return {
      registerKey: config.key, registerName: config.name, itemId: Number(row.ID || row.Id || 0), assetId: get(row, config.assetId),
      name: get(row, config.nameFields) || textValue(row.Title) || '(Unnamed asset)', status: get(row, config.status),
      assignedTo: get(row, config.assignedTo), custodian: get(row, config.custodian), location: get(row, config.location), warrantyStatus: get(row, config.warrantyStatus),
      expiry: expiryConfig ? get(row, expiryConfig.fields) : '', expiryType: expiryConfig?.type || '',
      totalCount: config.totalCount && fieldFor(config.totalCount) ? Number(get(row, config.totalCount)) || 0 : undefined
    };
  });
  return { config, assets, fields: selectedFields, hasStatus: !!fieldFor(config.status), hasQuantity: !!fieldFor(config.totalCount), hasExpiry: !!(config.expiry || []).find(item => fieldFor(item.fields)), warning };
};

const dashboardCache = new Map<string, { expiresAt: number; value: IDashboardRegisterData }>();
export const clearDashboardCache = (): void => dashboardCache.clear();
export const loadDashboardData = async (client: SPHttpClient, force = false, concurrency = 3): Promise<IDashboardData> => {
  const now = Date.now();
  const output: IDashboardRegisterData[] = new Array(dashboardRegisters.length);
  let cursor = 0;
  const worker = async (): Promise<void> => {
    while (true) {
      const index = cursor++;
      if (index >= dashboardRegisters.length) return;
      const config = dashboardRegisters[index];
      const cached = dashboardCache.get(config.key);
      if (!force && cached && cached.expiresAt > now) { output[index] = cached.value; continue; }
      try {
        const value = await metadataAndItems(client, config);
        dashboardCache.set(config.key, { expiresAt: now + 5 * 60 * 1000, value }); output[index] = value;
      } catch (error) { dashboardCache.delete(config.key); output[index] = { config, assets: [], fields: [], hasStatus: false, hasQuantity: false, hasExpiry: false, error: error instanceof Error ? error.message : 'Unable to load this register.' }; }
    }
  };
  await Promise.all(Array.from({ length: Math.max(1, Math.min(concurrency, dashboardRegisters.length)) }, () => worker()));
  return { registers: output, updatedAt: new Date() };
};

const localDate = (value: string | Date): Date | null => {
  if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  const match = value.trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
};
export const calculateDaysLeft = (expiry: string | Date, today = new Date()): number | null => {
  const date = localDate(expiry); if (!date) return null;
  const midnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return Math.ceil((Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - Date.UTC(midnight.getFullYear(), midnight.getMonth(), midnight.getDate())) / 86400000);
};
export const expiryPillTone = (daysLeft: number): ExpiryPillTone => daysLeft <= 7 ? 'red' : daysLeft <= 15 ? 'amber' : 'blue';
export const countStatuses = (assets: IDashboardAsset[]): Record<string, number> => {
  const result: Record<string, number> = { Allocated: 0, 'In Stock': 0, 'In Maintenance': 0, Retired: 0, Unspecified: 0 };
  assets.forEach(asset => {
    const status = asset.status.trim().toLowerCase();
    const canonical: Record<string, string> = { allocated: 'Allocated', 'in stock': 'In Stock', 'in maintenance': 'In Maintenance', retired: 'Retired' };
    const key = canonical[status] || (status ? 'Other' : 'Unspecified'); result[key] = (result[key] || 0) + 1;
  });
  return result;
};
export const expiringAssets = (assets: IDashboardAsset[], days = 30, today = new Date()): Array<IDashboardAsset & { daysLeft: number }> => assets.map(asset => ({ asset, daysLeft: calculateDaysLeft(asset.expiry, today) })).filter((item): item is { asset: IDashboardAsset; daysLeft: number } => item.daysLeft !== null && item.daysLeft >= 0 && item.daysLeft <= days).map(item => ({ ...item.asset, daysLeft: item.daysLeft })).sort((a, b) => a.daysLeft - b.daysLeft);
export const expiredAssets = (assets: IDashboardAsset[], today = new Date()): IDashboardAsset[] => assets.filter(asset => {
  const days = calculateDaysLeft(asset.expiry, today);
  return days !== null ? days < 0 : asset.warrantyStatus.trim().toLowerCase() === 'expired';
});
export const stockLevel = (inStock: number, total: number, minimum = 2): StockLevel => inStock === 0 || (total > 0 && inStock <= total * 0.1) ? 'Very low' : inStock <= minimum ? 'Low' : 'OK';
export const isAssigned = (value: string): boolean => !!value.trim() && ['unassigned', 'none', 'n/a', 'na'].indexOf(value.trim().toLowerCase()) < 0;
export interface IDashboardSummary { total: number; statuses: Record<string, number>; expiring: Array<IDashboardAsset & { daysLeft: number }>; expired: IDashboardAsset[]; noExpiryDate: number; duplicateIds: Array<{ id: string; count: number }>; byRegister: Array<{ config: IDashboardRegisterConfig; assets: IDashboardAsset[]; statuses: Record<string, number> }>; stocks: Array<{ config: IDashboardRegisterConfig; inStock: number; allocated: number; total: number; level: StockLevel }>; }
export const calculateDashboardSummary = (data: IDashboardData, windowDays = 30, today = new Date()): IDashboardSummary => {
  const successful = data.registers.filter(item => !item.error);
  const assets = successful.reduce<IDashboardAsset[]>((all, item) => all.concat(item.assets), []);
  const statusAssets = successful.filter(item => item.hasStatus).reduce<IDashboardAsset[]>((all, item) => all.concat(item.assets), []);
  const statuses = countStatuses(statusAssets);
  const expirables = successful.filter(item => item.hasExpiry).reduce<IDashboardAsset[]>((all, item) => all.concat(item.assets), []);
  const ids = new Map<string, number>(); assets.forEach(asset => { if (asset.assetId) ids.set(asset.assetId, (ids.get(asset.assetId) || 0) + 1); });
  const stocks = successful.filter(item => item.hasStatus || item.hasQuantity).map(item => {
    const quantity = item.hasQuantity && item.assets.some(asset => asset.totalCount !== undefined);
    const total = quantity ? item.assets.reduce((sum, asset) => sum + (asset.totalCount || 0), 0) : item.assets.length;
    const allocated = quantity ? item.assets.filter(asset => isAssigned(asset.assignedTo) || asset.status.trim().toLowerCase() === 'allocated').reduce((sum, asset) => sum + (asset.totalCount || 0), 0) : item.assets.filter(asset => asset.status.trim().toLowerCase() === 'allocated').length;
    const inStock = quantity ? Math.max(total - allocated, 0) : item.assets.filter(asset => asset.status.trim().toLowerCase() === 'in stock').length;
    return { config: item.config, inStock, allocated, total, level: stockLevel(inStock, total, dashboardStockThresholds[item.config.key] ?? 2) };
  });
  return {
    total: assets.length, statuses, expiring: expiringAssets(expirables, windowDays, today), expired: expiredAssets(expirables, today),
    noExpiryDate: expirables.filter(asset => !asset.expiry).length,
    duplicateIds: Array.from(ids.entries()).filter(([, count]) => count > 1).map(([id, count]) => ({ id, count })).sort((a, b) => b.count - a.count),
    byRegister: successful.map(item => ({ config: item.config, assets: item.assets, statuses: item.hasStatus ? countStatuses(item.assets) : {} })).sort((a, b) => b.assets.length - a.assets.length), stocks
  };
};
