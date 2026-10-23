import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { SPHttpClient } from '@microsoft/sp-http';
import styles from './AssetManagement.module.scss';
import AssetTimeline from './AssetTimeline';
import { buildAssetTimelineEvents, loadAssetTimelineSources, type IAssetTimelineSources } from './assetTimelineService';
import { getNextLaptopViewTab, laptopViewTabKeys, type LaptopViewTabKey } from './laptopViewTabNavigation';

type TabKey = LaptopViewTabKey;
interface IProps {
  itemId: number;
  purchaseDate: string;
  expiryDate: string;
  client: SPHttpClient;
  details: React.ReactNode;
  attachments: React.ReactNode;
  attachmentCount: number;
}
const tabs: Array<{ key: TabKey; title: string }> = [
  { key: 'details', title: 'Details' }, { key: 'timeline', title: 'Timeline' },
  { key: 'attachments', title: 'Attachments' }, { key: 'service', title: 'Service history' }
];
const panelId = (key: TabKey): string => `laptop-view-panel-${key}`;
const tabId = (key: TabKey): string => `laptop-view-tab-${key}`;
const sampleServiceHistory = (itemId: number): IAssetTimelineSources['requests'] => {
  const dateString = (monthsAgo: number, day: number): string => {
    const date = new Date();
    date.setMonth(date.getMonth() - monthsAgo, day);
    return `${date.getFullYear()}-${('0' + (date.getMonth() + 1)).slice(-2)}-${('0' + date.getDate()).slice(-2)}`;
  };
  return [
    { assetItemId: itemId, title: 'Sample: Battery failure', reportedDate: dateString(4, 12), resolvedDate: dateString(4, 15), technician: 'Alex Morgan', status: 'Closed' },
    { assetItemId: itemId, title: 'Sample: Keyboard replacement', reportedDate: dateString(2, 8), resolvedDate: dateString(2, 10), technician: 'Jordan Lee', status: 'Resolved' }
  ];
};

const LaptopViewTabs: React.FC<IProps> = ({ itemId, purchaseDate, expiryDate, client, details, attachments, attachmentCount }) => {
  const [activeTab, setActiveTab] = useState<TabKey>('details');
  const [sources, setSources] = useState<IAssetTimelineSources | null>(null);
  const [loading, setLoading] = useState(true);
  const [requestNumber, setRequestNumber] = useState(0);
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const refresh = async (): Promise<void> => {
    setLoading(true);
    const loaded = await loadAssetTimelineSources(client);
    // There is no service request list configured yet. Keep sample records scoped to
    // the selected SharePoint item and visibly label them so they are never mistaken
    // for live service history.
    setSources(loaded.requestSourceMissing ? { ...loaded, requests: sampleServiceHistory(itemId) } : loaded);
    setLoading(false);
  };
  useEffect(() => { void refresh(); }, [client, itemId, requestNumber]);
  const events = useMemo(() => sources ? buildAssetTimelineEvents(
    { itemId, purchaseDate, expiryDate }, sources.tracker, sources.requests
  ) : [], [itemId, purchaseDate, expiryDate, sources]);
  const onTabKeyDown = (event: React.KeyboardEvent<HTMLButtonElement>, key: TabKey): void => {
    const nextTab = getNextLaptopViewTab(key, event.key);
    if (!nextTab) return;
    event.preventDefault();
    setActiveTab(nextTab);
    tabRefs.current[laptopViewTabKeys.indexOf(nextTab)]?.focus();
  };
  const retry = (): void => setRequestNumber(value => value + 1);
  const sourceErrors = sources ? [sources.trackerError && `Asset Tracker: ${sources.trackerError}`, sources.requestError && `Service requests: ${sources.requestError}`].filter(Boolean) as string[] : [];
  return <div className={styles.laptopViewContent}>
    <div className={styles.viewTabs} role="tablist" aria-label="Laptop asset details">
      {tabs.map((tab, index) => <button ref={element => { tabRefs.current[index] = element; }} key={tab.key} id={tabId(tab.key)} type="button" role="tab" aria-selected={activeTab === tab.key} aria-controls={panelId(tab.key)} tabIndex={activeTab === tab.key ? 0 : -1} className={activeTab === tab.key ? styles.viewTabSelected : ''} onClick={() => setActiveTab(tab.key)} onKeyDown={event => onTabKeyDown(event, tab.key)}>{tab.title}{tab.key === 'attachments' ? ` (${attachmentCount})` : tab.key === 'service' ? ` (${sources?.requests.filter(item => item.assetItemId === itemId).length || 0})` : ''}</button>)}
    </div>
    <div className={styles.viewTabPanel} role="tabpanel" id={panelId('details')} aria-labelledby={tabId('details')} hidden={activeTab !== 'details'}>{details}</div>
    <div className={styles.viewTabPanel} role="tabpanel" id={panelId('timeline')} aria-labelledby={tabId('timeline')} hidden={activeTab !== 'timeline'}>
      {loading && !sources ? <div className={styles.timelineSkeleton} role="status" aria-label="Loading asset timeline"><i/><i/><i/></div> : <>
        {sources?.trackerError && <p className={styles.timelineNote} role="alert">Asset tracker history could not be loaded. <button type="button" onClick={retry}>Retry</button></p>}
        {sources?.requestError && <p className={styles.timelineNote} role="alert">Service history could not be loaded. <button type="button" onClick={retry}>Retry</button></p>}
        {sources?.requestSourceMissing && <p className={styles.timelineNote}>Showing sample service history. Configure a SharePoint service request source to show live records.</p>}
        <AssetTimeline events={events}/>
      </>}
    </div>
    <div className={styles.viewTabPanel} role="tabpanel" id={panelId('attachments')} aria-labelledby={tabId('attachments')} hidden={activeTab !== 'attachments'}>{attachments}</div>
    <div className={styles.viewTabPanel} role="tabpanel" id={panelId('service')} aria-labelledby={tabId('service')} hidden={activeTab !== 'service'}>
      {loading && !sources ? <div className={styles.timelineSkeleton} role="status" aria-label="Loading service history"><i/><i/><i/></div> : <>
        {sources?.requestError && <p className={styles.timelineNote} role="alert">Service history could not be loaded. <button type="button" onClick={retry}>Retry</button></p>}
        {sources?.requestSourceMissing && <p className={styles.timelineNote}>Showing sample service history. Configure a SharePoint service request source to show live records.</p>}
        {!!sources?.requests.length && <div className={styles.historyList}>{sources.requests.filter(item => item.assetItemId === itemId).map((record, index) => <article className={styles.historyItem} key={`${record.assetItemId}-${record.reportedDate}-${index}`}><div><strong>{record.title}</strong>{record.status && <span className={styles.historyStatus}>{record.status}</span>}</div><p>{[record.reportedDate && `Reported ${record.reportedDate}`, record.resolvedDate && `Resolved ${record.resolvedDate}`, record.technician && `Technician: ${record.technician}`].filter(Boolean).join(' · ')}</p></article>)}</div>}
        {sources && !sources.requests.some(item => item.assetItemId === itemId) && !sources.requestSourceMissing && !sources.requestError && <p className={styles.timelineEmpty}>No service history records found.</p>}
      </>}
    </div>
    {!!sourceErrors.length && <span className={styles.srOnly}>{sourceErrors.join('. ')}</span>}
  </div>;
};

export default LaptopViewTabs;
