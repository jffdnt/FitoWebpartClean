import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
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
    render(): void;
    protected onInit(): Promise<void>;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=WpCustomCoPilotWebPart.d.ts.map