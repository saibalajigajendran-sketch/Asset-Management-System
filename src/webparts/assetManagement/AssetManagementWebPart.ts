import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'AssetManagementWebPartStrings';
import AssetManagement from './components/AssetManagement';
import { IAssetManagementProps } from './components/IAssetManagementProps';
import { ensureFullScreenStyles, fullScreenBodyClass } from './components/globalStyles';

export interface IAssetManagementWebPartProps {
  description: string;
  hideChrome?: boolean;
}

export default class AssetManagementWebPart extends BaseClientSideWebPart<IAssetManagementWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _graphClient?: MSGraphClientV3;

  public render(): void {
    const isFullScreen = !this.context.sdks.microsoftTeams &&
      this.displayMode === DisplayMode.Read && this.properties.hideChrome !== false;
    if (isFullScreen) {
      ensureFullScreenStyles();
    }
    document.body.classList.toggle(fullScreenBodyClass, isFullScreen);

    const element: React.ReactElement<IAssetManagementProps> = React.createElement(
      AssetManagement,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName || this.context.pageContext.user.loginName,
        isFullScreen,
        spHttpClient: this.context.spHttpClient,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        graphClient: this._graphClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.all([
      this._getEnvironmentMessage(),
      this.context.msGraphClientFactory.getClient('3').catch(error => {
        console.error('Microsoft Graph client initialization failed', error);
        return undefined;
      })
    ]).then(([message, graphClient]) => {
      this._environmentMessage = message;
      this._graphClient = graphClient;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    document.body.classList.remove(fullScreenBodyClass);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneToggle('hideChrome', {
                  label: strings.FullScreenFieldLabel,
                  onText: strings.FullScreenOn,
                  offText: strings.FullScreenOff
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
