import { SPHttpClient, type MSGraphClientV3 } from '@microsoft/sp-http';

export interface ISharePointField {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
  Required: boolean;
  Choices?: string[] | { results: string[] };
}

export interface IEntraEmployee {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
}

interface ISharePointUser {
  Id: number;
  Title: string;
  Email?: string;
  LoginName?: string;
}

export type SharePointUserNameMap = Record<string, string>;
export interface ISharePointItemProjection { select: string; expand: string; }

interface ISharePointError {
  error?: { code?: string; message?: string | { value?: string } };
  'odata.error'?: { code?: string; message?: string | { value?: string } };
  message?: string;
}

interface IJsonResponse {
  ok: boolean;
  status: number;
  statusText?: string;
  text(): Promise<string>;
}

const listTitle = 'AssetRegister_Hardware_Laptop';
const acceptHeader = 'application/json;odata=nometadata';
const createHeaders = {
  Accept: acceptHeader,
  'Content-Type': 'application/json;odata=nometadata; charset=utf-8'
};

const normalize = (value: string): string => value
  .replace(/_x([0-9a-f]{4})_/gi, (_match, hex: string) => String.fromCharCode(parseInt(hex, 16)))
  .toLowerCase()
  .replace(/[^a-z0-9]/g, '');

const hiddenEditableFieldsByList: Record<string, string[]> = {
  assetregister_cloud_providers: ['cloudplatform', 'purpose'],
  assetregister_hardware_accessories: ['serialnumber'],
  assetregister_mediadevices: ['serialnumber'],
  assetregister_mobile_devices: ['assetowner', 'assetcustodian', 'maximumos'],
  assetregister_client_accessories: ['serialnumber', 'maximumos']
};

export const shouldIncludeSharePointFormField = (listUrl: string, field: ISharePointField): boolean => {
  const listName = listUrl.match(/\/Lists\/([^/]+)/i)?.[1].toLowerCase();
  const explicitFields = listName ? hiddenEditableFieldsByList[listName] || [] : [];
  const explicitlyRequested = explicitFields.indexOf(normalize(field.InternalName)) >= 0 || explicitFields.indexOf(normalize(field.Title)) >= 0;
  const systemFields = ['title', 'attachments', 'contenttype', 'contenttypeid'];
  return !field.ReadOnlyField && systemFields.indexOf(normalize(field.InternalName)) < 0 && (!field.Hidden || explicitlyRequested);
};

const readJson = async <T>(response: IJsonResponse, requestUrl?: string): Promise<T> => {
  const rawBody = await response.text();
  let body: T | ISharePointError = {} as ISharePointError;
  try {
    body = rawBody ? JSON.parse(rawBody) as T | ISharePointError : {} as ISharePointError;
  } catch {
    if (response.ok) throw new Error('SharePoint returned an unreadable response.');
  }
  if (!response.ok) {
    const error = body as ISharePointError;
    const errorMessage = (value?: string | { value?: string }): string | undefined =>
      typeof value === 'string' ? value : value?.value;
    const detail = errorMessage(error.error?.message) || errorMessage(error['odata.error']?.message) || error.message || rawBody;
    const code = error.error?.code || error['odata.error']?.code;
    if (requestUrl) console.error('SharePoint request failed', { url: requestUrl, status: response.status, body: error });
    throw new Error(`${detail || `SharePoint request failed (${response.status} ${response.statusText || ''}).`.trim()}${code ? ` [${code}]` : ''}`);
  }
  return body as T;
};

const readCollection = <T>(body: unknown, description: string): T[] => {
  const envelope = body as {
    d?: { results?: T[]; value?: T[] } | T[];
    value?: T[];
  };
  if (Array.isArray(envelope.d)) return envelope.d;
  if (envelope.d && !Array.isArray(envelope.d)) {
    if (Array.isArray(envelope.d.results)) return envelope.d.results;
    if (Array.isArray(envelope.d.value)) return envelope.d.value;
  }
  if (Array.isArray(envelope.value)) return envelope.value;
  throw new Error(`SharePoint returned an unexpected response while reading ${description}. Check list permissions and the SharePoint REST response.`);
};

export const buildSharePointItemProjection = (fields: ISharePointField[], selectedFields?: string[]): ISharePointItemProjection => {
  const userFields = fields.filter(field => field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti');
  if (!selectedFields) {
    const selectPeople = userFields.reduce<string[]>((all, field) => all.concat([`${field.InternalName}/Id`, `${field.InternalName}/Title`, `${field.InternalName}/EMail`]), []);
    return { select: `*,Attachments${selectPeople.length ? `,${selectPeople.join(',')}` : ''}`, expand: userFields.map(field => field.InternalName).join(',') };
  }
  const selectedUserFields = userFields.filter(field => selectedFields.indexOf(field.InternalName) >= 0);
  const selectedLookupFields = fields.filter(field => (field.TypeAsString === 'Lookup' || field.TypeAsString === 'LookupMulti') && selectedFields.indexOf(field.InternalName) >= 0);
  const peopleProperties = selectedUserFields.reduce<string[]>((all, field) => all.concat([`${field.InternalName}/Title`, `${field.InternalName}/EMail`]), []);
  const lookupProperties = selectedLookupFields.reduce<string[]>((all, field) => all.concat([`${field.InternalName}/Id`, `${field.InternalName}/Title`]), []);
  const expandedNames = Array.from(new Set([...selectedUserFields, ...selectedLookupFields].map(field => field.InternalName)));
  const expandedFields = expandedNames;
  const select = Array.from(new Set(['ID', 'Title', 'Attachments', ...selectedFields.filter(name => !expandedFields.some(field => field === name)), ...peopleProperties, ...lookupProperties])).join(',');
  return { select, expand: expandedNames.join(',') };
};

const resolveUserId = async (httpClient: SPHttpClient, webUrl: string, email: string): Promise<number> => {
  const response = await httpClient.get(
    `${webUrl}/_api/web/siteusers?$select=Id,Title,Email,LoginName&$top=5000`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: acceptHeader } }
  );
  const result = await readJson<unknown>(response);
  const users = readCollection<ISharePointUser>(result, 'site users');
  const emailKey = email.toLowerCase();
  const user = users.find(item => {
    const emailAddress = item.Email?.toLowerCase();
    const loginName = item.LoginName?.toLowerCase();
    const loginAlias = loginName?.split('|').pop();
    return emailAddress === emailKey || loginName === emailKey || loginAlias === emailKey;
  });
  if (!user) {
    throw new Error(`Could not resolve "${email}" to a SharePoint user on this site. Check that the employee has access to the site.`);
  }
  return user.Id;
};

export const getEntraEmployees = async (graphClient?: MSGraphClientV3): Promise<IEntraEmployee[]> => {
  if (!graphClient || typeof graphClient.api !== 'function') {
    throw new Error('The Microsoft Graph client is unavailable. Reload the web part before using the People Picker.');
  }
  const employees: IEntraEmployee[] = [];
  let page = await graphClient.api('/users')
    .version('v1.0')
    .select('id,displayName,mail,userPrincipalName')
    .top(999)
    .get() as { value: IEntraEmployee[]; '@odata.nextLink'?: string };

  employees.push(...page.value);
  while (page['@odata.nextLink']) {
    page = await graphClient.api(page['@odata.nextLink']).get() as { value: IEntraEmployee[]; '@odata.nextLink'?: string };
    employees.push(...page.value);
  }

  return employees
    .filter(employee => !!employee.displayName && !!(employee.mail || employee.userPrincipalName))
    .sort((left, right) => left.displayName.localeCompare(right.displayName, undefined, { sensitivity: 'base' }));
};

const saveListItemWithMetadata = async (
  httpClient: SPHttpClient,
  webUrl: string,
  listUrl: string,
  fields: ISharePointField[],
  values: Record<string, string | number>,
  itemId?: number
): Promise<number | undefined> => {
  const payload: Record<string, unknown> = {};
  const submittedFields: string[] = [];

  const labels = Object.keys(values);
  for (let index = 0; index < labels.length; index += 1) {
    const label = labels[index];
    const rawValue = values[label];
    if (rawValue === '') continue;
    const labelKey = normalize(label);
    let field = fields.filter(item => normalize(item.Title) === labelKey || normalize(item.InternalName) === labelKey)[0];
    if (!field && labelKey === 'name') {
      field = fields.filter(item => item.InternalName === 'Title')[0];
    }
    if (!field) {
      if (labelKey === 'operatingsystemdetails') {
        throw new Error(`Could not match OperatingSystemDetails to a writable SharePoint column in "${listTitle}". Check its exact column title/internal name.`);
      }
      continue;
    }

    const fieldType = field.TypeAsString;
    if (fieldType === 'User' || fieldType === 'UserMulti') {
      const submittedPeople = fieldType === 'UserMulti' ? String(rawValue).split(';#').filter(Boolean) : [String(rawValue)];
      const userIds: number[] = [];
      for (let userIndex = 0; userIndex < submittedPeople.length; userIndex += 1) {
        const submittedPerson = submittedPeople[userIndex];
        userIds.push(/^\d+$/.test(submittedPerson)
          ? Number(submittedPerson)
          : await resolveUserId(httpClient, webUrl, submittedPerson));
      }
      payload[`${field.InternalName}Id`] = fieldType === 'UserMulti' ? { results: userIds } : userIds[0];
      submittedFields.push(field.InternalName);
    } else if (fieldType === 'Lookup' || fieldType === 'LookupMulti') {
      if (fieldType === 'LookupMulti' || !/^\d+$/.test(String(rawValue))) continue;
      payload[`${field.InternalName}Id`] = Number(rawValue);
      submittedFields.push(field.InternalName);
    } else if (fieldType === 'Number' || fieldType === 'Currency') {
      const numericValue = Number(rawValue);
      if (!Number.isFinite(numericValue)) {
        throw new Error(`Enter a numeric value for SharePoint field "${field.Title}".`);
      }
      payload[field.InternalName] = numericValue;
      submittedFields.push(field.InternalName);
    } else if (fieldType === 'DateTime') {
      const dateValue = new Date(String(rawValue));
      if (!Number.isFinite(dateValue.getTime())) {
        throw new Error(`Enter a valid date for "${field.Title}".`);
      }
      payload[field.InternalName] = dateValue.toISOString();
      submittedFields.push(field.InternalName);
    } else if (fieldType === 'Choice') {
      const choices = Array.isArray(field.Choices) ? field.Choices : field.Choices?.results || [];
      const selectedChoice = choices.filter(choice => choice.toLowerCase() === String(rawValue).toLowerCase())[0];
      if (!selectedChoice) {
        throw new Error(`Invalid value for "${field.Title}". Select one of: ${choices.join(', ')}.`);
      }
      payload[field.InternalName] = selectedChoice;
      submittedFields.push(field.InternalName);
    } else if (fieldType === 'MultiChoice') {
      const choices = Array.isArray(field.Choices) ? field.Choices : field.Choices?.results || [];
      const selectedChoices = String(rawValue).split(';#').filter(Boolean).map(value => {
        const choice = choices.filter(candidate => candidate.toLowerCase() === value.toLowerCase())[0];
        if (!choice) throw new Error(`Invalid value for "${field.Title}". Select from: ${choices.join(', ')}.`);
        return choice;
      });
      payload[field.InternalName] = { results: selectedChoices };
      submittedFields.push(field.InternalName);
    } else if (fieldType === 'Boolean') {
      payload[field.InternalName] = String(rawValue).toLowerCase() === 'true';
      submittedFields.push(field.InternalName);
    } else {
      payload[field.InternalName] = String(rawValue);
      submittedFields.push(field.InternalName);
    }
  }

  const missingRequired = fields.filter(field => field.Required && submittedFields.indexOf(field.InternalName) < 0);
  if (missingRequired.length) {
    throw new Error(`Complete required SharePoint fields before saving: ${missingRequired.map(field => field.Title).join(', ')}.`);
  }

  const requestHeaders = itemId === undefined
    ? { ...createHeaders, Prefer: 'return=representation' }
    : { ...createHeaders, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' };
  const itemsUrl = listUrl.replace(/\?@listUrl=(.*)$/, '/items?@listUrl=$1');
  const saveResponse = await httpClient.post(
    itemId === undefined ? itemsUrl : itemsUrl.replace('/items?', `/items(${itemId})?`),
    SPHttpClient.configurations.v1,
    { headers: requestHeaders, body: JSON.stringify(payload) }
  );
  const saved = await readJson<unknown>(saveResponse) as { ID?: number; Id?: number; d?: { ID?: number; Id?: number }; __metadata?: { id?: string } };
  const metadataId = saved.__metadata?.id?.match(/items\((\d+)\)/i)?.[1];
  return itemId === undefined ? Number(saved.ID || saved.Id || saved.d?.ID || saved.d?.Id || metadataId) || undefined : itemId;
};

export const getSharePointListFields = async (
  httpClient: SPHttpClient,
  listUrl: string
): Promise<ISharePointField[]> => {
  const match = listUrl.match(/^https?:\/\/[^/]+(\/.*)$/i);
  if (!match) throw new Error('The configured SharePoint list URL is invalid.');
  const serverRelativeListUrl = decodeURIComponent(match[1].split('/AllItems.aspx')[0]).replace(/'/g, "''");
  const webUrl = listUrl.match(/^(https?:\/\/[^/]+)/i)?.[1];
  if (!webUrl) throw new Error('The configured SharePoint list URL is invalid.');
  const requestUrl = `${webUrl}${serverRelativeListUrl.substring(0, serverRelativeListUrl.indexOf('/Lists/'))}/_api/web/GetList(@listUrl)/fields?@listUrl='${encodeURIComponent(serverRelativeListUrl)}'&$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField,Required,Choices`;
  const response = await httpClient.get(requestUrl, SPHttpClient.configurations.v1, { headers: { Accept: acceptHeader } });
  const result = await readJson<unknown>(response, requestUrl);
  return readCollection<ISharePointField>(result, 'list fields').filter(field => shouldIncludeSharePointFormField(listUrl, field));
};

export const saveSharePointListItem = (
  httpClient: SPHttpClient,
  webUrl: string,
  listUrl: string,
  fields: ISharePointField[],
  values: Record<string, string | number>,
  itemId?: number
): Promise<number | undefined> => {
  const listMatch = listUrl.match(/^https?:\/\/[^/]+(\/.*)$/i);
  if (!listMatch) throw new Error('The configured SharePoint list URL is invalid.');
  const serverRelativeListUrl = decodeURIComponent(listMatch[1].split('/AllItems.aspx')[0]).replace(/'/g, "''");
  const listWebUrl = listUrl.match(/^(https?:\/\/[^/]+)/i)?.[1];
  if (!listWebUrl) throw new Error('The configured SharePoint list URL is invalid.');
  const baseListUrl = `${listWebUrl}${serverRelativeListUrl.substring(0, serverRelativeListUrl.indexOf('/Lists/'))}/_api/web/GetList(@listUrl)?@listUrl='${encodeURIComponent(serverRelativeListUrl)}'`;
  return saveListItemWithMetadata(httpClient, webUrl, baseListUrl, fields, values, itemId);
};

export const deleteSharePointListItem = async (
  httpClient: SPHttpClient,
  listUrl: string,
  itemId: number
): Promise<void> => {
  const listMatch = listUrl.match(/^https?:\/\/[^/]+(\/.*)$/i);
  if (!listMatch || !Number.isInteger(itemId) || itemId <= 0) throw new Error('Cannot delete this SharePoint item because its list URL or item ID is invalid.');
  const serverRelativeListUrl = decodeURIComponent(listMatch[1].split('/AllItems.aspx')[0]).replace(/'/g, "''");
  const listWebUrl = listUrl.match(/^(https?:\/\/[^/]+)/i)?.[1];
  if (!listWebUrl) throw new Error('The configured SharePoint list URL is invalid.');
  const itemsUrl = `${listWebUrl}${serverRelativeListUrl.substring(0, serverRelativeListUrl.indexOf('/Lists/'))}/_api/web/GetList(@listUrl)/items(${itemId})?@listUrl='${encodeURIComponent(serverRelativeListUrl)}'`;
  const response = await httpClient.post(itemsUrl, SPHttpClient.configurations.v1, {
    headers: { ...createHeaders, 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' },
    body: ''
  });
  await readJson<unknown>(response);
};

interface ISharePointItemsPage<T> {
  value?: T[];
  'odata.nextLink'?: string;
  '@odata.nextLink'?: string;
  d?: { results?: T[]; __next?: string };
}

export const getSharePointListItems = async (
  httpClient: SPHttpClient,
  listUrl: string,
  fields: ISharePointField[] = [],
  selectedFields?: string[]
): Promise<Record<string, unknown>[]> => {
  const match = listUrl.match(/^https?:\/\/[^/]+(\/.*)$/i);
  if (!match) throw new Error('The configured SharePoint list URL is invalid.');
  const serverRelativeListUrl = decodeURIComponent(match[1].split('/AllItems.aspx')[0]).replace(/'/g, "''");
  const sitePath = serverRelativeListUrl.substring(0, serverRelativeListUrl.indexOf('/Lists/'));
  const webUrl = listUrl.match(/^(https?:\/\/[^/]+)/i)?.[1];
  if (!webUrl) throw new Error('The configured SharePoint list URL is invalid.');
  const projection = buildSharePointItemProjection(fields, selectedFields);
  const expandQuery = projection.expand ? `&$select=${projection.select}&$expand=${projection.expand}` : `&$select=${projection.select}`;
  const requestUrl = `${webUrl}${sitePath}/_api/web/GetList(@listUrl)/items?@listUrl='${encodeURIComponent(serverRelativeListUrl)}'&$top=5000${expandQuery}`;
  const rows: Record<string, unknown>[] = [];
  let nextUrl: string | undefined = requestUrl;

  while (nextUrl) {
    const response = await httpClient.get(nextUrl, SPHttpClient.configurations.v1, { headers: { Accept: acceptHeader } });
    const result: ISharePointItemsPage<Record<string, unknown>> = await readJson<ISharePointItemsPage<Record<string, unknown>>>(response, nextUrl);
    rows.push(...readCollection<Record<string, unknown>>(result, 'list items'));
    nextUrl = result['@odata.nextLink'] || result['odata.nextLink'] || result.d?.__next;
  }

  return rows;
};

export const getSharePointUserNames = async (httpClient: SPHttpClient, webUrl: string): Promise<SharePointUserNameMap> => {
  const response = await httpClient.get(
    `${webUrl}/_api/web/siteusers?$select=Id,Title&$top=5000`,
    SPHttpClient.configurations.v1,
    { headers: { Accept: acceptHeader } }
  );
  const result = await readJson<unknown>(response);
  const users = readCollection<ISharePointUser>(result, 'SharePoint site users');
  const names: SharePointUserNameMap = {};
  users.forEach(user => { names[String(user.Id)] = user.Title; });
  return names;
};
