import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'InboundCustomerV2WebPartStrings';
import InboundCustomerV2 from './components/InboundCustomerV2';
import { InboundCustomerV2Props } from './Interfaces/InboundCustomerV2Props';

export interface InboundCustomerV2WebPartProps {
  description: string;
}

export default class InboundCustomerV2WebPart extends BaseClientSideWebPart<InboundCustomerV2Props> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<InboundCustomerV2Props> = React.createElement(
      InboundCustomerV2,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        hubSiteUrl: this.properties.hubSiteUrl,
        projectInformationListName: this.properties.projectInformationListName,
        hubsite: this.properties.hubsite,
        TransmittalIDSettings: this.properties.TransmittalIDSettings,
        InboundTransmittalHeader: this.properties.InboundTransmittalHeader,
        InboundTransmittalDetails: this.properties.InboundTransmittalDetails,
        OutboundTransmittalHeader: this.properties.OutboundTransmittalHeader,
        OutboundTransmittalDetails: this.properties.OutboundTransmittalDetails,
        InboundAdditionalDocuments: this.properties.InboundAdditionalDocuments,
        documentIndexList: this.properties.documentIndexList,
        TransmittalHistory: this.properties.TransmittalHistory,
        TransmittalOutlookLibrary: this.properties.TransmittalOutlookLibrary,
        TransmittalCodeSettings: this.properties.TransmittalCodeSettings,
        EmailNotificationSettings: this.properties.EmailNotificationSettings,
        NotificationPreferenceSettings: this.properties.NotificationPreferenceSettings,
        requestList: this.properties.requestList,
        redirectUrl: this.properties.redirectUrl,
        PermissionMatrixSettings: this.properties.PermissionMatrixSettings,
        accessGroupDetailsList:this.properties.accessGroupDetailsList
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
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
                PropertyPaneTextField('hubSiteUrl', {
                  label: "hubSiteUrl"
                }),
                PropertyPaneTextField('projectInformationListName', {
                  label: "projectInformationListName"
                }),
                PropertyPaneTextField('hubsite', {
                  label: "hubsite"
                }),
                PropertyPaneTextField('TransmittalIDSettings', {
                  label: "TransmittalIDSettings"
                }),
                PropertyPaneTextField('InboundTransmittalHeader', {
                  label: "InboundTransmittalHeader"
                }),
                PropertyPaneTextField('InboundTransmittalDetails', {
                  label: "InboundTransmittalDetails"
                }),
                PropertyPaneTextField('OutboundTransmittalHeader', {
                  label: "OutboundTransmittalHeader"
                }),
                PropertyPaneTextField('OutboundTransmittalDetails', {
                  label: "OutboundTransmittalDetails"
                }),
                PropertyPaneTextField('InboundAdditionalDocuments',{
                  label: "InboundAdditionalDocuments"
                }),
                PropertyPaneTextField('documentIndexList', {
                  label: "documentIndexList"
                }),
                PropertyPaneTextField('TransmittalHistory', {
                  label: "TransmittalHistory"
                }),
                PropertyPaneTextField('PermissionMatrixSettings', {
                  label: "PermissionMatrixSettings"
                }),
                PropertyPaneTextField('TransmittalCodeSettings', {
                  label: "TransmittalCodeSettings"
                }),
                PropertyPaneTextField('TransmittalOutlookLibrary', {
                  label: "TransmittalOutlookLibrary"
                }),
                PropertyPaneTextField('EmailNotificationSettings', {
                  label: "EmailNotificationSettings"
                }),
                PropertyPaneTextField('NotificationPreferenceSettings', {
                  label: "NotificationPreferenceSettings"
                }),
                PropertyPaneTextField('requestList', {
                  label: "Request"
                }),
                PropertyPaneTextField('redirectUrl', {
                  label: "redirectUrl"
                }),
                PropertyPaneTextField('accessGroupDetailsList', {
                  label: "accessGroupDetailsList"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
