import { SPHttpClient } from '@microsoft/sp-http';
import { getSharePointListFields, getSharePointListItems, type ISharePointField } from './hardwareLaptopSharePointService';
import { assetTrackerTimelineSource, serviceRequestTimelineSource, type ITimelineSourceConfig } from './assetTimelineConfig';

export type TimelineEventType = 'purchased' | 'movement' | 'failure' | 'resolved' | 'support-expiring' | 'support-expired' | 'support-valid';
export interface IAssetTimelineEvent { id: string; date: Date; title: string; description: string; type: TimelineEventType; position: 'above' | 'below'; }
export interface IAssetTimelineAsset { itemId: number; assetId?: string; purchaseDate: string; expiryDate: string; }
export interface ITimelineTrackerEvent { assetItemId: number; date: string; from: string; to: string; approvedBy: string; }
export interface ITimelineServiceRequest { assetItemId: number; title: string; reportedDate: string; resolvedDate: string; technician: string; status: string; }
export interface IAssetTimelineSources { tracker: ITimelineTrackerEvent[]; requests: ITimelineServiceRequest[]; trackerError?: string; requestError?: string; requestSourceMissing?: boolean; trackerFields: string[]; requestFields: string[]; }

const normalize = (value: string): string => value.replace(/_x([0-9a-f]{4})_/gi, (_match, hex: string) => String.fromCharCode(parseInt(hex, 16))).toLowerCase().replace(/[^a-z0-9]/g, '');
const dateAtLocalMidnight = (value: string | Date): Date | undefined => {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? undefined : new Date(value.getFullYear(), value.getMonth(), value.getDate());
  const trimmed = value.trim();
  const match = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  const parsed = match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])) : new Date(trimmed);
  const date = parsed;
  return Number.isNaN(date.getTime()) ? undefined : new Date(date.getFullYear(), date.getMonth(), date.getDate());
};
export const calculateTimelineDaysLeft = (expiry: string | Date, today = new Date()): number | null => {
  const date = dateAtLocalMidnight(expiry); if (!date) return null;
  const midnight = dateAtLocalMidnight(today) as Date;
  return Math.ceil((Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - Date.UTC(midnight.getFullYear(), midnight.getMonth(), midnight.getDate())) / 86400000);
};
const readText = (value: unknown): string => {
  if (value === null || value === undefined) return '';
  if (typeof value !== 'object') return String(value).trim();
  const record = value as { Title?: string; LookupValue?: string; EMail?: string; Email?: string; results?: Array<{ Title?: string; LookupValue?: string }>; value?: Array<{ Title?: string; LookupValue?: string }>; d?: { results?: Array<{ Title?: string; LookupValue?: string }> } };
  return record.Title || record.LookupValue || record.EMail || record.Email || (record.results || record.value || record.d?.results)?.map(item => item.Title || item.LookupValue || '').filter(Boolean).join('; ') || '';
};
const valueFor = (row: Record<string, unknown>, fields: ISharePointField[], aliases: string[] = []): unknown => {
  for (const alias of aliases) {
    const field = fields.find(candidate => normalize(candidate.Title) === normalize(alias) || normalize(candidate.InternalName) === normalize(alias));
    if (field && Object.prototype.hasOwnProperty.call(row, field.InternalName)) return row[field.InternalName];
  }
  for (const alias of aliases) {
    const key = Object.keys(row).find(candidate => normalize(candidate) === normalize(alias));
    if (key) return row[key];
  }
  return undefined;
};
const referenceId = (value: unknown): number | undefined => {
  if (typeof value === 'number' && Number.isInteger(value) && value > 0) return value;
  if (typeof value === 'string' && /^\d+$/.test(value.trim())) return Number(value);
  if (value && typeof value === 'object') {
    const record = value as { Id?: number | string; ID?: number | string; LookupId?: number | string; results?: Array<{ Id?: number | string; ID?: number | string; LookupId?: number | string }> };
    if (record.results?.length) return referenceId(record.results[0]);
    const id = Number(record.Id || record.ID || record.LookupId);
    return Number.isInteger(id) && id > 0 ? id : undefined;
  }
  return undefined;
};
const fieldByAliases = (fields: ISharePointField[], aliases: string[]): ISharePointField | undefined => fields.find(field => aliases.some(alias => normalize(alias) === normalize(field.Title) || normalize(alias) === normalize(field.InternalName)));
const selected = (fields: ISharePointField[], aliases: string[]): string[] => Array.from(new Set(aliases.map(alias => fieldByAliases(fields, [alias])?.InternalName).filter((name): name is string => !!name)));
const loadSourceRows = async (client: SPHttpClient, source: ITimelineSourceConfig): Promise<{ rows: Record<string, unknown>[]; fields: ISharePointField[]; selectedFields: string[] }> => {
  const fields = await getSharePointListFields(client, source.listUrl);
  const referenceFields = source.assetItemIdFields.filter(alias => {
    const field = fieldByAliases(fields, [alias]);
    if (!field) return false;
    return field.TypeAsString === 'Lookup' || field.TypeAsString === 'LookupMulti' || /itemid|recordid/i.test(field.InternalName);
  });
  const selectedFields = selected(fields, [...referenceFields, ...source.dateFields, ...(source.reportedFields || []), ...(source.resolvedFields || []), ...(source.titleFields || []), ...(source.descriptionFields || []), ...(source.fromFields || []), ...(source.toFields || []), ...(source.approvedByFields || []), ...(source.technicianFields || []), ...(source.statusFields || [])]);
  return { rows: await getSharePointListItems(client, source.listUrl, fields, selectedFields), fields, selectedFields };
};
const mappedSource = async (client: SPHttpClient, source: ITimelineSourceConfig, tracker: boolean): Promise<{ rows: Array<ITimelineTrackerEvent | ITimelineServiceRequest>; fields: string[] }> => {
  const loaded = await loadSourceRows(client, source);
  const refAliases = source.assetItemIdFields.filter(alias => {
    const field = fieldByAliases(loaded.fields, [alias]);
    return !!field && (field.TypeAsString === 'Lookup' || field.TypeAsString === 'LookupMulti' || /itemid|recordid/i.test(field.InternalName));
  });
  if (!refAliases.length) throw new Error(`No SharePoint item-ID reference field was found. Configure the correct relationship field for ${tracker ? 'Asset Tracker' : 'service requests'}.`);
  const rows = loaded.rows.map(row => {
    const assetItemId = referenceId(valueFor(row, loaded.fields, refAliases));
    if (!assetItemId) return undefined;
    if (tracker) return { assetItemId, date: readText(valueFor(row, loaded.fields, source.dateFields)), from: readText(valueFor(row, loaded.fields, source.fromFields)), to: readText(valueFor(row, loaded.fields, source.toFields || source.titleFields)), approvedBy: readText(valueFor(row, loaded.fields, source.approvedByFields)) } as ITimelineTrackerEvent;
    return { assetItemId, title: readText(valueFor(row, loaded.fields, source.titleFields)) || 'Service request', reportedDate: readText(valueFor(row, loaded.fields, source.reportedFields || source.dateFields)), resolvedDate: readText(valueFor(row, loaded.fields, source.resolvedFields)), technician: readText(valueFor(row, loaded.fields, source.technicianFields)), status: readText(valueFor(row, loaded.fields, source.statusFields)) } as ITimelineServiceRequest;
  }).filter((row): row is ITimelineTrackerEvent | ITimelineServiceRequest => !!row);
  return { rows, fields: loaded.selectedFields };
};
export const loadAssetTimelineSources = async (client: SPHttpClient): Promise<IAssetTimelineSources> => {
  const trackerPromise = mappedSource(client, assetTrackerTimelineSource, true);
  const requestPromise = serviceRequestTimelineSource ? mappedSource(client, serviceRequestTimelineSource, false) : undefined;
  const settle = <T>(promise: Promise<T>): Promise<{ value: T; error?: undefined } | { value?: undefined; error: unknown }> => promise.then(value => ({ value }), error => ({ error }));
  const [trackerResult, requestResult] = await Promise.all([settle(trackerPromise), settle(requestPromise || Promise.resolve({ rows: [], fields: [] }))]);
  const tracker = trackerResult.value || { rows: [] as Array<ITimelineTrackerEvent | ITimelineServiceRequest>, fields: [] as string[] };
  const requests = requestResult.value || { rows: [] as Array<ITimelineTrackerEvent | ITimelineServiceRequest>, fields: [] as string[] };
  return {
    tracker: tracker.rows.filter((row): row is ITimelineTrackerEvent => 'date' in row && 'from' in row),
    requests: requests.rows.filter((row): row is ITimelineServiceRequest => 'reportedDate' in row),
    trackerFields: tracker.fields, requestFields: requests.fields,
    trackerError: trackerResult.error ? trackerResult.error instanceof Error ? trackerResult.error.message : 'Unable to load Asset tracker history.' : undefined,
    requestError: requestResult.error ? requestResult.error instanceof Error ? requestResult.error.message : 'Unable to load service requests.' : undefined,
    requestSourceMissing: !serviceRequestTimelineSource
  };
};

export const buildAssetTimelineEvents = (asset: IAssetTimelineAsset, trackerRows: ITimelineTrackerEvent[], requests: ITimelineServiceRequest[], today = new Date()): IAssetTimelineEvent[] => {
  const events: Array<Omit<IAssetTimelineEvent, 'position'>> = [];
  const add = (id: string, dateValue: string | Date, title: string, description: string, type: TimelineEventType): void => {
    const date = dateAtLocalMidnight(dateValue);
    if (date) events.push({ id, date, title, description, type });
  };
  add('purchased', asset.purchaseDate, 'Purchased & Installed', 'Asset purchased and installed.', 'purchased');
  trackerRows.filter(row => row.assetItemId === asset.itemId).forEach((row, index) => {
    const title = row.to || 'Asset moved';
    const parts = [row.from ? `Moved from ${row.from}` : '', row.to ? `to ${row.to}` : '', row.approvedBy ? `approved by ${row.approvedBy}` : ''].filter(Boolean);
    add(`movement-${index}`, row.date, title, parts.join(', ') || 'Asset movement recorded.', 'movement');
  });
  requests.filter(row => row.assetItemId === asset.itemId).forEach((row, index) => {
    const technician = row.technician ? `Technician: ${row.technician}` : '';
    add(`failure-${index}`, row.reportedDate, row.title, technician || 'Service issue reported.', 'failure');
    if (/resolved|closed/i.test(row.status)) add(`resolved-${index}`, row.resolvedDate, `${row.title} resolved`, technician || 'Service request resolved.', 'resolved');
  });
  const days = calculateTimelineDaysLeft(asset.expiryDate, today);
  if (days !== null) {
    if (days < 0) add('support', asset.expiryDate, 'Support expired', 'Renewal mail sent to Asset Owner and Asset Custodian.', 'support-expired');
    else if (days <= 30) add('support', asset.expiryDate, 'Support expiring', `${days} ${days === 1 ? 'day' : 'days'} left.`, 'support-expiring');
    else add('support', asset.expiryDate, `Support valid until ${dateAtLocalMidnight(asset.expiryDate)?.toLocaleDateString() || ''}`, 'Support coverage is currently valid.', 'support-valid');
  }
  return events.sort((a, b) => a.date.getTime() - b.date.getTime()).map((event, index) => ({ ...event, position: index % 2 === 0 ? 'above' : 'below' }));
};
export const selectInitialTimelineEvents = (events: IAssetTimelineEvent[], limit = 8): IAssetTimelineEvent[] => {
  if (events.length <= limit) return events;
  const keep = new Set<string>();
  const purchase = events.find(event => event.type === 'purchased');
  const support = events.find(event => event.type.indexOf('support-') === 0);
  if (purchase) keep.add(purchase.id);
  if (support) keep.add(support.id);
  for (let index = events.length - 1; index >= 0 && keep.size < limit; index -= 1) keep.add(events[index].id);
  return events.filter(event => keep.has(event.id)).map((event, index) => ({ ...event, position: index % 2 === 0 ? 'above' : 'below' }));
};
