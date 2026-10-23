jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));

import { buildAssetTimelineEvents, calculateTimelineDaysLeft, selectInitialTimelineEvents, type IAssetTimelineAsset, type IAssetTimelineEvent, type ITimelineServiceRequest, type ITimelineTrackerEvent } from './assetTimelineService';
import { getNextLaptopViewTab } from './laptopViewTabNavigation';

const asset = (overrides: Partial<IAssetTimelineAsset> = {}): IAssetTimelineAsset => ({ itemId: 14, assetId: 'DUPLICATE-ID', purchaseDate: '2026-01-01', expiryDate: '', ...overrides });
const event = (id: string, type: IAssetTimelineEvent['type'], date: string, position: IAssetTimelineEvent['position'] = 'above'): IAssetTimelineEvent => ({ id, type, date: new Date(`${date}T00:00:00`), title: id, description: id, position });

describe('asset timeline', () => {
  it('builds and chronologically sorts purchase, movement, failure, resolution and support events', () => {
    const tracker: ITimelineTrackerEvent[] = [{ assetItemId: 14, date: '2026-03-20', from: 'Storage', to: 'User A', approvedBy: 'Manager B' }];
    const requests: ITimelineServiceRequest[] = [{ assetItemId: 14, title: 'Battery failure', reportedDate: '2026-03-20', resolvedDate: '2026-03-25', technician: 'Tech C', status: 'Closed' }];
    const events = buildAssetTimelineEvents(asset({ expiryDate: '2026-03-30' }), tracker, requests, new Date(2026, 2, 1));
    expect(events.map(item => item.type)).toEqual(['purchased', 'movement', 'failure', 'resolved', 'support-expiring']);
    expect(events[1].description).toBe('Moved from Storage, to User A, approved by Manager B');
    expect(events[2].description).toContain('Tech C');
  });

  it('keeps chronological order within a month and alternates event card positions', () => {
    const events = buildAssetTimelineEvents(asset({ purchaseDate: '2026-06-01' }), [
      { assetItemId: 14, date: '2026-06-20', from: 'A', to: 'B', approvedBy: '' },
      { assetItemId: 14, date: '2026-06-08', from: 'B', to: 'C', approvedBy: '' }
    ], [], new Date(2026, 0, 1));
    expect(events.map(item => item.date.getDate())).toEqual([1, 8, 20]);
    expect(events.map(item => item.position)).toEqual(['above', 'below', 'above']);
  });

  it('classifies expired and expiring support at the 0, 1, 30 and 31 day boundaries', () => {
    const today = new Date(2026, 0, 1);
    const at = (days: number): string => { const date = new Date(2026, 0, 1 + days); return `${date.getFullYear()}-${('0' + (date.getMonth() + 1)).slice(-2)}-${('0' + date.getDate()).slice(-2)}`; };
    expect(buildAssetTimelineEvents(asset({ expiryDate: at(-1) }), [], [], today).find(item => item.type.startsWith('support-'))?.type).toBe('support-expired');
    [0, 1, 30].forEach(days => expect(buildAssetTimelineEvents(asset({ expiryDate: at(days) }), [], [], today).find(item => item.type.startsWith('support-'))?.type).toBe('support-expiring'));
    expect(buildAssetTimelineEvents(asset({ expiryDate: at(31) }), [], [], today).find(item => item.type.startsWith('support-'))?.type).toBe('support-valid');
    expect(calculateTimelineDaysLeft(at(30), today)).toBe(30);
  });

  it('skips missing and invalid dates', () => {
    const events = buildAssetTimelineEvents(asset({ purchaseDate: 'not-a-date', expiryDate: '' }), [
      { assetItemId: 14, date: '', from: '', to: 'Desk', approvedBy: '' }
    ], [{ assetItemId: 14, title: 'Failure', reportedDate: 'invalid', resolvedDate: '', technician: '', status: 'Open' }]);
    expect(events).toHaveLength(0);
  });

  it('limits a long history while retaining purchase, support, and the newest events', () => {
    const events = [event('purchase', 'purchased', '2020-01-01'), ...Array.from({ length: 10 }, (_, index) => event(`move-${index}`, 'movement', `2021-01-${('0' + (index + 1)).slice(-2)}`)), event('support', 'support-expiring', '2026-01-01')];
    const visible = selectInitialTimelineEvents(events);
    expect(visible).toHaveLength(8);
    expect(visible[0].id).toBe('purchase');
    expect(visible.some(item => item.id === 'support')).toBe(true);
    expect(visible[visible.length - 1].id).toBe('support');
    expect(visible.map(item => item.position)).toEqual(['above', 'below', 'above', 'below', 'above', 'below', 'above', 'below']);
  });

  it('associates tracker and service events by SharePoint item ID when Asset IDs repeat', () => {
    const tracker: ITimelineTrackerEvent[] = [
      { assetItemId: 14, date: '2026-02-01', from: 'A', to: 'B', approvedBy: '' },
      { assetItemId: 25, date: '2026-02-02', from: 'C', to: 'D', approvedBy: '' }
    ];
    const requests: ITimelineServiceRequest[] = [
      { assetItemId: 14, title: 'Screen failure', reportedDate: '2026-02-03', resolvedDate: '', technician: '', status: 'Open' },
      { assetItemId: 25, title: 'Keyboard failure', reportedDate: '2026-02-04', resolvedDate: '', technician: '', status: 'Open' }
    ];
    const first = buildAssetTimelineEvents(asset({ itemId: 14 }), tracker, requests);
    const second = buildAssetTimelineEvents(asset({ itemId: 25 }), tracker, requests);
    expect(first.some(item => item.title === 'Screen failure')).toBe(true);
    expect(first.some(item => item.title === 'Keyboard failure')).toBe(false);
    expect(second.some(item => item.title === 'Keyboard failure')).toBe(true);
    expect(second.some(item => item.title === 'Screen failure')).toBe(false);
  });

  it('supports arrow, Home and End keyboard navigation across view tabs', () => {
    expect(getNextLaptopViewTab('details', 'ArrowRight')).toBe('timeline');
    expect(getNextLaptopViewTab('details', 'ArrowLeft')).toBe('service');
    expect(getNextLaptopViewTab('attachments', 'Home')).toBe('details');
    expect(getNextLaptopViewTab('timeline', 'End')).toBe('service');
    expect(getNextLaptopViewTab('timeline', 'Enter')).toBeUndefined();
  });
});
