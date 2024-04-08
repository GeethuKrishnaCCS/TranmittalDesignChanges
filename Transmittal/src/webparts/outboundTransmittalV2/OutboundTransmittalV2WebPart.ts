import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'OutboundTransmittalV2WebPartStrings';
import OutboundTransmittalV2 from './components/OutboundTransmittalV2';
import { IOutboundTransmittalV2Props } from './Interfaces/IOutboundTransmittalV2Props';

export interface IOutboundTransmittalV2WebPartProps {
  description: string;
}

export default class OutboundTransmittalV2WebPart extends BaseClientSideWebPart<IOutboundTransmittalV2Props> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IOutboundTransmittalV2Props> = React.createElement(
      OutboundTransmittalV2,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        projectInformationListName: this.properties.projectInformationListName,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        companiesListName: this.properties.companiesListName,
        hubSiteUrl: this.properties.hubSiteUrl,
        hubSite: this.properties.hubSite,
        contactListName: this.properties.contactListName,
        publishDocumentLibraryName: this.properties.publishDocumentLibraryName,
        transmittalCodeSettingsListName: this.properties.transmittalCodeSettingsListName,
        sourceDocumentLibraryName: this.properties.sourceDocumentLibraryName,
        maxFileSize: this.properties.maxFileSize,
        transmittalIdSettingsListName: this.properties.transmittalIdSettingsListName,
        outboundTransmittalHeaderListName: this.properties.outboundTransmittalHeaderListName,
        outBoundTransmittalSitePage: this.properties.outBoundTransmittalSitePage,
        outboundTransmittalDetailsListName: this.properties.outboundTransmittalDetailsListName,
        outboundAdditionalDocumentsListName: this.properties.outboundAdditionalDocumentsListName,
        userMessageSettings: this.properties.userMessageSettings,
        transmittalHistoryLogList: this.properties.transmittalHistoryLogList,
        notificationPrefListName: this.properties.notificationPrefListName,
        emailNotificationSettings: this.properties.emailNotificationSettings,
        documentIndex: this.properties.documentIndex,
        masterListName: this.properties.masterListName,
        modalBGColor: this.properties.modalBGColor,
        permissionMatrixSettings: this.properties.permissionMatrixSettings,
        accessGroupDetailsListName: this.properties.accessGroupDetailsListName
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
                PropertyPaneTextField('projectInformationListName', {
                  label: "Project Information List Name"
                }),
                PropertyPaneTextField('hubSite', {
                  label: "Hub Site Name"
                }),
                PropertyPaneTextField('hubSiteUrl', {
                  label: "Hub Site Url"
                }),
                PropertyPaneTextField('companiesListName', {
                  label: "Companies List Name"
                }),
                PropertyPaneTextField('contactListName', {
                  label: "Contact List Name"
                }),
                PropertyPaneTextField('publishDocumentLibraryName', {
                  label: "Published Document Library Name"
                }),
                PropertyPaneTextField('transmittalCodeSettingsListName', {
                  label: "Transmittal Code Settings List Name"
                }),
                PropertyPaneTextField('sourceDocumentLibraryName', {
                  label: "SourceDocument Library Name"
                }),
                PropertyPaneTextField('transmittalIdSettingsListName', {
                  label: "Transmittal Id Settings List Name"
                }),
                PropertyPaneTextField('outboundTransmittalHeaderListName', {
                  label: "Outbound Transmittal Header List Name"
                }),
                PropertyPaneTextField('outboundTransmittalDetailsListName', {
                  label: "Outbound Transmittal Details List Name"
                }),
                PropertyPaneTextField('outboundAdditionalDocumentsListName', {
                  label: "Outbound Additional Documents List Name"
                }),
                PropertyPaneTextField('outBoundTransmittalSitePage', {
                  label: "Outbound Transmittal Page Name"
                }),
                PropertyPaneTextField('userMessageSettings', {
                  label: "User Message Settings list Name"
                }),
                PropertyPaneTextField('emailNotificationSettings', {
                  label: "Email Notification Settings List Name"
                }),
                PropertyPaneTextField('transmittalHistoryLogList', {
                  label: "Transmittal History Log list Name"
                }),
                PropertyPaneTextField('notificationPrefListName', {
                  label: "Notification Preference List Name"
                }),
                PropertyPaneTextField('masterListName', {
                  label: "Master List Name"
                }),
                PropertyPaneTextField('maxFileSize', {
                  label: "Max File Size"
                }),
                PropertyPaneTextField('modalBGColor', {
                  label: "Preview BG Color"
                }),
                PropertyPaneTextField('permissionMatrixSettings', {
                  label: "Permission Matrix Settings List Name"
                }),
                PropertyPaneTextField('accessGroupDetailsListName', {
                  label: "Access Group Details List Name"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
