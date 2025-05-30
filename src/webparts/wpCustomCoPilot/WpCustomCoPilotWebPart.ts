import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import WpCustomCoPilot from './components/WpCustomCoPilot';
import { IWpCustomCoPilotProps } from './components/IWpCustomCoPilotProps';

export interface IWpCustomCoPilotWebPartProps {
  botName: string;
  botURL: string;
  clientID: string;
  authority: string;
  customScope: string;
  greet: boolean;
  userDisplayName: string;
  userEmail: string;
  userFriendlyName: string;
  welcomeMessage: string;
  botAvatarImage: string;
  botAvatarInitials: string;
  height?: string;
  width?: string;
  headerHeight?: string;
  headerBgColor?: string;
  headerTextColor?: string;
  headerFontSize?: string;
  chatContainerPaddingTop?: string;
  headerPaddingLeft?: string;
}

export default class WpCustomCoPilotWebPart extends BaseClientSideWebPart<IWpCustomCoPilotWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWpCustomCoPilotProps> = React.createElement(
      WpCustomCoPilot,
      {
        botName: this.properties.botName,
        botURL: this.properties.botURL,
        clientID: this.properties.clientID,
        authority: this.properties.authority,
        customScope: this.properties.customScope,
        greet: this.properties.greet,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        userFriendlyName: this.context.pageContext.user.displayName,
        welcomeMessage: this.properties.welcomeMessage,
        botAvatarImage: this.properties.botAvatarImage,
        botAvatarInitials: this.properties.botAvatarInitials,
        height: this.properties.height,
        width: this.properties.width,
        headerHeight: this.properties.headerHeight,
        headerBgColor: this.properties.headerBgColor,
        headerTextColor: this.properties.headerTextColor,
        headerFontSize: this.properties.headerFontSize,
        chatContainerPaddingTop: this.properties.chatContainerPaddingTop,
        headerPaddingLeft: this.properties.headerPaddingLeft
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    return Promise.resolve();
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
          header: { description: 'Configure your Copilot Web Part' },
          groups: [
            {
              groupName: 'Bot Settings',
              groupFields: [
                PropertyPaneTextField('botName', { label: 'Bot Name' }),
                PropertyPaneTextField('botURL', { label: 'Bot URL' }),
                PropertyPaneTextField('clientID', { label: 'Client ID' }),
                PropertyPaneTextField('authority', { label: 'Authority' }),
                PropertyPaneTextField('customScope', { label: 'Custom Scope' }),
                PropertyPaneToggle('greet', { label: 'Greet User' }),
                PropertyPaneTextField('welcomeMessage', { label: 'Welcome Message' }),
                PropertyPaneTextField('botAvatarImage', { label: 'Bot Avatar Image' }),
                PropertyPaneTextField('botAvatarInitials', { label: 'Bot Avatar Initials' }),
                PropertyPaneTextField('height', { label: 'Chat Height (px)', description: 'e.g. 400' }),
                PropertyPaneTextField('width', { label: 'Chat Width (px)', description: 'e.g. 100%' }),
                PropertyPaneTextField('headerHeight', { label: 'Header Height (px)', description: 'e.g. 60' }),
                PropertyPaneTextField('headerBgColor', { label: 'Header Background Color', description: 'e.g. #009FDB' }),
                PropertyPaneTextField('headerTextColor', { label: 'Header Text Color', description: 'e.g. #fff' }),
                PropertyPaneTextField('headerFontSize', { label: 'Header Font Size (px)', description: 'e.g. 20' }),
                PropertyPaneTextField('chatContainerPaddingTop', { label: 'Chat Padding Top (px)', description: 'e.g. 16' }),
                PropertyPaneTextField('headerPaddingLeft', { label: 'Header Padding Left (px)', description: 'e.g. 24' })
              ]
            }
          ]
        }
      ]
    };
  }
}
