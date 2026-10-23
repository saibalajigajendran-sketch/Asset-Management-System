import { SPHttpClient } from '@microsoft/sp-http';

export interface IAssetAttachment { FileName: string; ServerRelativeUrl: string; Length?: string | number; }
export const allowedExtensions = ['pdf', 'jpg', 'jpeg', 'png', 'docx', 'xlsx', 'msg', 'eml'];
export const maxAttachmentSize = 10 * 1024 * 1024;
export const maxSharePointPathLength = 400;

export const validateAttachmentFiles = (files: File[], alreadyStaged: File[] = [], existingNames: string[] = [], attachmentPathPrefix = ''): string[] => {
  const errors: string[] = [];
  if (files.length + alreadyStaged.length > 10) errors.push('A maximum of 10 files can be uploaded in one batch.');
  const names = new Set(existingNames.concat(alreadyStaged.map(file => file.name)).map(name => name.toLowerCase()));
  files.forEach(file => {
    const extension = file.name.split('.').pop()?.toLowerCase() || '';
    if (!file.size) errors.push(`${file.name}: empty files are not allowed.`);
    else if (file.size > maxAttachmentSize) errors.push(`${file.name}: files must be 10 MB or smaller.`);
    if (allowedExtensions.indexOf(extension) < 0) errors.push(`${file.name}: unsupported file type.`);
    if (names.has(file.name.toLowerCase())) errors.push(`A file named "${file.name}" already exists on this asset. Rename the file or delete the existing one.`);
    names.add(file.name.toLowerCase());
    const illegal = Array.from(new Set(file.name.match(/[#%&*<>?:\/\\{|}~]/g) || []));
    if (illegal.length) {
      illegal.forEach(character => errors.push(`${file.name}: This file name contains "${character}", which SharePoint does not allow. Please rename the file and upload it again.`));
    }
    if (file.name.startsWith('.') || file.name.startsWith(' ') || file.name.endsWith('.') || file.name.endsWith(' ')) {
      errors.push(`${file.name}: SharePoint file names cannot start or end with a dot or space. Please rename the file and upload it again.`);
    }
    if (Array.from(`${attachmentPathPrefix}${file.name}`).length > maxSharePointPathLength) {
      errors.push(`${file.name}: The file path exceeds SharePoint's ${maxSharePointPathLength}-character limit. Please rename the file and upload it again.`);
    }
  });
  return errors;
};

export const uploadFilesWithResults = async <T>(files: T[], upload: (file: T) => Promise<void>, onStart?: (file: T) => void): Promise<{ succeeded: T[]; failed: T[] }> => {
  const succeeded: T[] = []; const failed: T[] = [];
  for (const file of files) { try { if (onStart) onStart(file); await upload(file); succeeded.push(file); } catch { failed.push(file); } }
  return { succeeded, failed };
};

export const getAttachmentPathPrefix = (listUrl: string, itemId: number): string => {
  const match = listUrl.match(/^https?:\/\/[^/]+(\/.*)$/i);
  if (!match) throw new Error('The configured SharePoint list URL is invalid.');
  let listPath = match[1].split('/AllItems.aspx')[0];
  try { listPath = decodeURIComponent(listPath); } catch { /* Keep the SharePoint path if it contains a literal percent. */ }
  return `${listPath}/Attachments/${itemId}/`;
};

const encodeODataName = (name: string): string => encodeURIComponent(name.replace(/'/g, "''"))
  .replace(/[!'()*]/g, character => `%${character.charCodeAt(0).toString(16).toUpperCase()}`);

const collection = (body: unknown): IAssetAttachment[] => {
  const data = body as { value?: IAssetAttachment[]; d?: { results?: IAssetAttachment[]; value?: IAssetAttachment[] } };
  return data.value || data.d?.results || data.d?.value || [];
};
const listBase = (webUrl: string, listUrl: string): { resource: string; listParameter: string } => {
  const match = listUrl.match(/^https?:\/\/[^/]+(\/.*)$/i);
  if (!match) throw new Error('The configured SharePoint list URL is invalid.');
  const path = decodeURIComponent(match[1].split('/AllItems.aspx')[0]).replace(/'/g, "''");
  const sitePath = path.substring(0, path.indexOf('/Lists/'));
  const origin = listUrl.match(/^(https?:\/\/[^/]+)/i)?.[1] || webUrl.match(/^(https?:\/\/[^/]+)/i)?.[1];
  if (!origin) throw new Error('The configured SharePoint web URL is invalid.');
  return { resource: `${origin}${sitePath}/_api/web/GetList(@listUrl)`, listParameter: `@listUrl='${encodeURIComponent(path)}'` };
};
const readResponse = async <T>(response: { ok: boolean; status: number; text(): Promise<string> }): Promise<T> => {
  const raw = await response.text();
  if (!response.ok) {
    if (response.status === 401 || response.status === 403) throw new Error('You do not have permission to manage attachments on this SharePoint list. Check your list permissions and sign-in status.');
    let detail = raw;
    try { const parsed = JSON.parse(raw) as { error?: { message?: string | { value?: string } } }; detail = String(parsed.error?.message && typeof parsed.error.message === 'object' ? parsed.error.message.value : parsed.error?.message || raw); } catch { /* SharePoint may return a plain-text error. */ }
    throw new Error(detail || `SharePoint request failed (${response.status}).`);
  }
  if (!raw) return {} as T;
  return JSON.parse(raw) as T;
};
export const checkAttachmentsEnabled = async (client: SPHttpClient, webUrl: string, listUrl: string): Promise<void> => {
  const base = listBase(webUrl, listUrl);
  const enabled = await readResponse<{ EnableAttachments?: boolean }>(await client.get(`${base.resource}?${base.listParameter}&$select=EnableAttachments`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } }));
  if (enabled.EnableAttachments === false) throw new Error('Attachments are disabled for this SharePoint list. Ask a SharePoint administrator to enable attachments for this register.');
};
export const listAttachments = async (client: SPHttpClient, webUrl: string, listUrl: string, itemId: number): Promise<IAssetAttachment[]> => {
  const base = listBase(webUrl, listUrl);
  await checkAttachmentsEnabled(client, webUrl, listUrl);
  const body = await readResponse<unknown>(await client.get(`${base.resource}/items(${itemId})/AttachmentFiles?${base.listParameter}&$select=FileName,ServerRelativeUrl,Length`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } }));
  return collection(body);
};
export const getAttachmentUrl = (webUrl: string, file: IAssetAttachment): string => {
  const origin = webUrl.match(/^(https?:\/\/[^/]+)/i)?.[1] || webUrl.replace(/\/$/, '');
  const path = file.ServerRelativeUrl.split('/').map(segment => {
    let decoded = segment;
    try { decoded = decodeURIComponent(segment); } catch { /* Encode literal URL characters as provided by SharePoint. */ }
    return encodeURIComponent(decoded).replace(/[!'()*]/g, character => `%${character.charCodeAt(0).toString(16).toUpperCase()}`);
  }).join('/');
  return `${origin}${path}`;
};
export const uploadAttachment = async (client: SPHttpClient, webUrl: string, listUrl: string, itemId: number, file: File): Promise<void> => {
  const base = listBase(webUrl, listUrl);
  const response = await client.post(`${base.resource}/items(${itemId})/AttachmentFiles/add(FileName='${encodeODataName(file.name)}')?${base.listParameter}`, SPHttpClient.configurations.v1, {
    headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/octet-stream' },
    body: await file.arrayBuffer()
  } as never);
  await readResponse<unknown>(response);
};
export const deleteAttachment = async (client: SPHttpClient, webUrl: string, listUrl: string, itemId: number, fileName: string): Promise<void> => {
  const base = listBase(webUrl, listUrl);
  const response = await client.post(`${base.resource}/items(${itemId})/AttachmentFiles/getByFileName('${encodeODataName(fileName)}')?${base.listParameter}`, SPHttpClient.configurations.v1, {
    headers: { Accept: 'application/json;odata=nometadata', 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }, body: ''
  });
  await readResponse<unknown>(response);
};
