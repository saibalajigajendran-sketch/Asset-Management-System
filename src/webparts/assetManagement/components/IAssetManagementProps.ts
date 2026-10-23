import type { MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';

export interface IAssetManagementProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  isFullScreen: boolean;
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  graphClient?: MSGraphClientV3;
}
