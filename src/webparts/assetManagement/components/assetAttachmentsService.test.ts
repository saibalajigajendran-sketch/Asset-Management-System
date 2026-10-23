jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));

import { getAttachmentUrl, getAttachmentPathPrefix, listAttachments, uploadAttachment, uploadFilesWithResults, validateAttachmentFiles, type IAssetAttachment } from './assetAttachmentsService';
import type { SPHttpClient } from '@microsoft/sp-http';

const listUrl = 'https://tenant.example/sites/assets/Lists/Laptops/AllItems.aspx';
const webUrl = 'https://tenant.example/sites/assets';
const attachmentClient = (filesByItem: Record<number, IAssetAttachment[]>): SPHttpClient => ({
  get: jest.fn(async (url: string) => {
    if (url.indexOf('$select=EnableAttachments') >= 0) return { ok: true, status: 200, text: async () => JSON.stringify({ EnableAttachments: true }) };
    const match = url.match(/items\((\d+)\)/);
    const files = filesByItem[Number(match?.[1])] || [];
    return { ok: true, status: 200, text: async () => JSON.stringify({ value: files }) };
  }),
  post: jest.fn(async () => ({ ok: true, status: 200, text: async () => '{}' }))
} as unknown as SPHttpClient);
const makeFile = (name: string, size = 10): File => ({ name, size, arrayBuffer: async () => new ArrayBuffer(size) } as File);

describe('attachment helpers', () => {
  it('loads files for the selected SharePoint item', async () => {
    const file: IAssetAttachment = { FileName: 'invoice.final.pdf', ServerRelativeUrl: '/sites/assets/Lists/Laptops/Attachments/21/invoice.final.pdf', Length: 1024 };
    const files = await listAttachments(attachmentClient({ 21: [file] }), webUrl, listUrl, 21);
    expect(files).toEqual([file]);
  });

  it('returns an empty list when the selected item has no files', async () => {
    await expect(listAttachments(attachmentClient({}), webUrl, listUrl, 21)).resolves.toEqual([]);
  });

  it('keeps attachments separate for records with the same Asset ID', async () => {
    const first: IAssetAttachment = { FileName: 'receipt.pdf', ServerRelativeUrl: '/attachments/21/receipt.pdf' };
    const second: IAssetAttachment = { FileName: 'receipt.pdf', ServerRelativeUrl: '/attachments/22/receipt.pdf' };
    const client = attachmentClient({ 21: [first], 22: [second] });
    expect(await listAttachments(client, webUrl, listUrl, 21)).toEqual([first]);
    expect(await listAttachments(client, webUrl, listUrl, 22)).toEqual([second]);
  });

  it('keeps original names with spaces, capitals, multiple dots, apostrophes, brackets, and unicode', async () => {
    const names = ['Quarterly Report.FINAL.PDF', 'User\'s [Résumé].v2.PDF'];
    for (const name of names) {
      const client = attachmentClient({});
      await uploadAttachment(client, webUrl, listUrl, 21, makeFile(name),);
      const requestUrl = (client.post as jest.Mock).mock.calls[0][0] as string;
      expect(decodeURIComponent(requestUrl)).toContain(name.replace(/'/g, "''"));
    }
    expect(getAttachmentUrl(webUrl, { FileName: names[1], ServerRelativeUrl: `/sites/assets/Lists/Laptops/Attachments/21/${encodeURIComponent(names[1])}` }))
      .toContain('User%27s%20%5BR%C3%A9sum%C3%A9%5D.v2.PDF');
  });

  it('rejects case-insensitive duplicates against existing attachments and staged files', () => {
    expect(validateAttachmentFiles([makeFile('REPORT.pdf')], [], ['report.PDF']))
      .toContain('A file named "REPORT.pdf" already exists on this asset. Rename the file or delete the existing one.');
    expect(validateAttachmentFiles([makeFile('report.pdf')], [makeFile('Report.PDF')]))
      .toContain('A file named "report.pdf" already exists on this asset. Rename the file or delete the existing one.');
  });

  it('rejects every SharePoint-invalid filename character with the offending character in the message', () => {
    for (const character of ['#', '%', '&', '*', ':', '<', '>', '?', '/', '\\', '{', '|', '}', '~']) {
      const errors = validateAttachmentFiles([makeFile(`bad${character}name.pdf`)]);
      expect(errors.some(error => error.includes(`contains "${character}"`))).toBe(true);
    }
  });

  it('rejects names starting or ending with a dot or space', () => {
    for (const name of ['.hidden.pdf', 'bad name.pdf ', ' trailing.pdf', 'trailing..']) {
      expect(validateAttachmentFiles([makeFile(name)]).some(error => error.includes('cannot start or end with a dot or space'))).toBe(true);
    }
  });

  it('rejects names whose full attachment path exceeds SharePoint path length', () => {
    const prefix = getAttachmentPathPrefix(listUrl, 9876543210);
    const longName = `${'x'.repeat(401 - prefix.length)}.pdf`;
    expect(validateAttachmentFiles([makeFile(longName)], [], [], prefix).some(error => error.includes('path exceeds SharePoint'))).toBe(true);
  });

  it('preserves allowed type, size, empty file, and batch-count validation', () => {
    expect(validateAttachmentFiles([makeFile('ok.pdf')])).toEqual([]);
    expect(validateAttachmentFiles([makeFile('empty.pdf', 0)])).toContain('empty.pdf: empty files are not allowed.');
    expect(validateAttachmentFiles([makeFile('bad.exe')])).toContain('bad.exe: unsupported file type.');
    expect(validateAttachmentFiles([makeFile('large.pdf', 10 * 1024 * 1024 + 1)])).toContain('large.pdf: files must be 10 MB or smaller.');
    expect(validateAttachmentFiles(Array.from({ length: 11 }, (_, i) => makeFile(`${i}.pdf`)))).toContain('A maximum of 10 files can be uploaded in one batch.');
  });

  it('uploads successfully when Cloud Providers has no identifier field', async () => {
    const client = attachmentClient({});
    await expect(uploadAttachment(client, webUrl, listUrl, 21, makeFile('cloud export.v1.xlsx'))).resolves.toBeUndefined();
    expect(client.post).toHaveBeenCalledTimes(1);
    expect(decodeURIComponent((client.post as jest.Mock).mock.calls[0][0] as string)).toContain("FileName='cloud export.v1.xlsx'");
  });

  it('keeps failed uploads for retry while allowing successful uploads to finish', async () => {
    const calls: string[] = [];
    const result = await uploadFilesWithResults(['one', 'two', 'three'], async file => {
      calls.push(file);
      if (file === 'two') throw new Error('network failure');
    });
    expect(calls).toEqual(['one', 'two', 'three']);
    expect(result.succeeded).toEqual(['one', 'three']);
    expect(result.failed).toEqual(['two']);
  });
});
