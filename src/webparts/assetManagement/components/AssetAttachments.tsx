import * as React from 'react';
import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';
import { Icon } from '@fluentui/react';
import styles from './AssetManagement.module.scss';
import { checkAttachmentsEnabled, deleteAttachment, getAttachmentPathPrefix, getAttachmentUrl, listAttachments, uploadAttachment, uploadFilesWithResults, validateAttachmentFiles, type IAssetAttachment } from './assetAttachmentsService';
import type { SPHttpClient } from '@microsoft/sp-http';

export interface IStagedAttachment { file: File; error?: string; progress?: string; }
export interface AssetAttachmentsHandle { uploadStaged(itemId: number): Promise<number>; hasStaged(): boolean; }
interface IProps { listUrl: string; webUrl: string; client: SPHttpClient; itemId?: number; mode: 'add' | 'edit' | 'view'; onAttachmentsChanged?: (hasAttachments: boolean) => void; onCountChange?: (count: number) => void; }

const AssetAttachments = forwardRef<AssetAttachmentsHandle, IProps>(({ listUrl, webUrl, client, itemId, mode, onAttachmentsChanged, onCountChange }, ref) => {
  const [attachments, setAttachments] = useState<IAssetAttachment[]>([]);
  const [staged, setStaged] = useState<IStagedAttachment[]>([]);
  const [errors, setErrors] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [confirm, setConfirm] = useState<IAssetAttachment | null>(null);
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const requestNumber = useRef(0);
  const refresh = async (): Promise<void> => {
    const currentRequest = ++requestNumber.current;
    setLoading(true); setErrors([]); onCountChange?.(0);
    try {
      if (itemId) {
        const files = await listAttachments(client, webUrl, listUrl, itemId);
        if (requestNumber.current === currentRequest) { setAttachments(files); onAttachmentsChanged?.(files.length > 0); onCountChange?.(files.length); }
      } else {
        await checkAttachmentsEnabled(client, webUrl, listUrl);
        if (requestNumber.current === currentRequest) { setAttachments([]); onCountChange?.(0); }
      }
    } catch (error) {
      if (requestNumber.current === currentRequest) setErrors([error instanceof Error ? error.message : 'Unable to load attachments.']);
    } finally { if (requestNumber.current === currentRequest) setLoading(false); }
  };
  useEffect(() => { void refresh(); return () => { requestNumber.current += 1; }; }, [itemId, listUrl]);
  const addFiles = (files: FileList | File[]): void => {
    const selected = Array.from(files);
    const accepted: File[] = [];
    const problems: string[] = [];
    selected.forEach(file => {
      if (staged.length + accepted.length >= 10) { problems.push(`${file.name}: a maximum of 10 files can be uploaded in one batch.`); return; }
      const fileErrors = validateAttachmentFiles([file], staged.map(item => item.file).concat(accepted), attachments.map(item => item.FileName), itemId ? getAttachmentPathPrefix(listUrl, itemId) : '');
      if (fileErrors.length) problems.push(...fileErrors); else accepted.push(file);
    });
    setErrors(problems);
    if (accepted.length) setStaged(current => [...current, ...accepted.map(file => ({ file }))]);
  };
  useImperativeHandle(ref, () => ({
    hasStaged: () => staged.length > 0,
    uploadStaged: async (newItemId: number): Promise<number> => {
      setBusy(true);
      let currentAttachments: IAssetAttachment[];
      try { currentAttachments = await listAttachments(client, webUrl, listUrl, newItemId); }
      catch (error) {
        const message = error instanceof Error ? error.message : 'Unable to check existing attachment names.';
        setStaged(current => current.map(item => ({ ...item, error: message })));
        setBusy(false);
        return staged.length;
      }
      const existing = currentAttachments.map(file => file.FileName);
      const pathPrefix = getAttachmentPathPrefix(listUrl, newItemId);
      const uploadErrors = new Map<IStagedAttachment, string>();
      const results = await uploadFilesWithResults(staged, async entry => {
          try {
            const validationErrors = validateAttachmentFiles([entry.file], [], existing, pathPrefix);
            if (validationErrors.length) throw new Error(validationErrors.join(' '));
            await uploadAttachment(client, webUrl, listUrl, newItemId, entry.file);
            existing.push(entry.file.name);
            setStaged(current => current.filter(item => item !== entry));
          } catch (error) {
            uploadErrors.set(entry, error instanceof Error ? error.message : 'Upload failed. Check your connection and SharePoint permissions, then retry.');
            throw error;
          }
      }, entry => setStaged(current => current.map(item => item === entry ? { ...item, progress: 'Uploading…', error: undefined } : item)));
      results.failed.forEach(entry => {
          const message = uploadErrors.get(entry) || 'Upload failed. Check your connection and SharePoint permissions, then retry.';
          setStaged(current => current.map(item => item === entry ? { ...item, progress: undefined, error: message } : item));
      });
      await refresh(); setBusy(false); return results.failed.length;
    }
  }), [staged, attachments, client, webUrl, listUrl]);
  const confirmDelete = async (): Promise<void> => {
    if (!confirm || !itemId) return;
    setBusy(true);
    try { await deleteAttachment(client, webUrl, listUrl, itemId, confirm.FileName); setConfirm(null); await refresh(); }
    catch (error) { setErrors([error instanceof Error ? error.message : 'Unable to delete attachment.']); }
    finally { setBusy(false); }
  };
  const openFile = (file: IAssetAttachment, download: boolean): void => {
    const url = getAttachmentUrl(webUrl, file);
    if (download) { const link = document.createElement('a'); link.href = url; link.download = file.FileName; link.click(); }
    else window.open(url, '_blank', 'noopener,noreferrer');
  };
  const previewable = (name: string): boolean => /\.(pdf|jpe?g|png)$/i.test(name);
  const editable = mode !== 'view';
  return <section className={styles.attachmentSection} aria-label="Attachments">
    <div className={styles.attachmentHeading}><div><h3>{mode === 'view' ? `Attachments (${attachments.length})` : 'Attachments'}</h3><small>Files are stored on this SharePoint list item.</small></div>{loading && <span className={styles.attachmentSkeleton} aria-label="Loading attachments" />}</div>
    {errors.map((error, index) => <p className={styles.attachmentError} role="alert" key={`${index}-${error}`}>{error}<button type="button" onClick={() => void refresh()}>Retry</button></p>)}
    {attachments.length === 0 && !loading && !errors.length && <p className={styles.attachmentEmpty}>{mode === 'view' ? 'No files yet.' : 'No files yet. Upload an invoice or warranty card.'}</p>}
    {attachments.map(file => <article className={styles.attachmentRow} key={file.FileName}><Icon iconName="Attach"/><span className={styles.attachmentName}><strong>{file.FileName}</strong>{file.Length !== undefined && file.Length !== '' && <small>{(Number(file.Length) / 1024).toFixed(0)} KB</small>}</span><div className={styles.attachmentActions}><button type="button" aria-label={`View ${file.FileName}`} onClick={() => openFile(file, !previewable(file.FileName))}>View</button><button type="button" aria-label={`Download ${file.FileName}`} onClick={() => openFile(file, true)}>Download</button>{editable && <button type="button" aria-label={`Delete ${file.FileName}`} onClick={() => setConfirm(file)}>Delete</button>}</div></article>)}
    {staged.map((entry, index) => <article className={styles.attachmentRow} key={`${entry.file.name}-${index}`}><Icon iconName="Attach"/><span className={styles.attachmentName}><strong>{entry.file.name}</strong><small>{(entry.file.size / 1024).toFixed(0)} KB {entry.error ? `· ${entry.error}` : `· ${entry.progress || 'Ready to upload'}`}</small></span><button type="button" aria-label={`Remove ${entry.file.name}`} disabled={busy} onClick={() => setStaged(current => current.filter(item => item !== entry))}>×</button></article>)}
    {editable && <div className={`${styles.attachmentDrop} ${dragging ? styles.attachmentDragging : ''}`} onDragOver={event => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={event => { event.preventDefault(); setDragging(false); addFiles(event.dataTransfer.files); }}><Icon iconName="CloudUpload"/><span>Drag files here or</span><button type="button" disabled={busy} onClick={() => inputRef.current?.click()}>Choose files</button><input ref={inputRef} type="file" multiple aria-label="Choose attachment files" accept=".pdf,.jpg,.jpeg,.png,.docx,.xlsx,.msg,.eml" onChange={event => { if (event.currentTarget.files) addFiles(event.currentTarget.files); event.currentTarget.value = ''; }}/><small>PDF, JPG, PNG, DOCX, XLSX, MSG, EML · up to 10 MB each · 10 files max</small></div>}
    {confirm && <div className={styles.confirmBackdrop} role="presentation"><section className={styles.confirmDialog} role="alertdialog" aria-modal="true"><h2>Delete attachment?</h2><p>Delete <strong>{confirm.FileName}</strong> from this asset?</p><footer className={styles.confirmActions}><button type="button" className={styles.secondaryButton} disabled={busy} onClick={() => setConfirm(null)}>Cancel</button><button type="button" className={styles.deleteButton} disabled={busy} onClick={() => void confirmDelete()}>{busy ? 'Deleting...' : 'Delete file'}</button></footer></section></div>}
  </section>;
});
AssetAttachments.displayName = 'AssetAttachments';
export default AssetAttachments;
