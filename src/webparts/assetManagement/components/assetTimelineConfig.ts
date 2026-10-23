import { assetRegisterUrls } from './dashboardConfig';

export interface ITimelineSourceConfig {
  listUrl: string;
  assetItemIdFields: string[];
  dateFields: string[];
  titleFields?: string[];
  descriptionFields?: string[];
  fromFields?: string[];
  toFields?: string[];
  approvedByFields?: string[];
  reportedFields?: string[];
  resolvedFields?: string[];
  technicianFields?: string[];
  statusFields?: string[];
}

export const assetTrackerTimelineSource: ITimelineSourceConfig = {
  listUrl: assetRegisterUrls['asset-tracker'],
  assetItemIdFields: ['AssetItemID', 'AssetItemId', 'AssetRecordID', 'AssetRecordId', 'AssetLookup', 'Asset'],
  dateFields: ['DateOfMovement', 'Date Of Movement'],
  titleFields: ['AssetMovedTo', 'Asset Moved To'],
  fromFields: ['AssetMovedFrom', 'Asset Moved From'],
  toFields: ['AssetMovedTo', 'Asset Moved To'],
  approvedByFields: ['ApprovedBy', 'Approved By', 'ApprovalBy']
};

// The existing app has no SharePoint service-request list configured. Set this
// to the configured list URL when that source is identified; no list is guessed.
export const serviceRequestTimelineSource: ITimelineSourceConfig | undefined = undefined;
