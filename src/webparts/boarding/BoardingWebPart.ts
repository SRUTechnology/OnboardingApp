import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'BoardingWebPartStrings';
import { ListEnsureResult ,Web} from "sp-pnp-js";
import Boarding from './components/Boarding';
import { IBoardingProps } from './components/IBoardingProps';
import ContextService from "./components/Services.ts/ContextService";

export interface IBoardingWebPartProps {
  description: string;
}

export default class BoardingWebPart extends BaseClientSideWebPart<IBoardingWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IBoardingProps> = React.createElement(
      Boarding,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // private listsToBeCreated = [
  //   {
  //     type: "list",
  //     listname: "EOSettings",
  //     Titles: ["FirstTimeInstallation"],
  //     xmlData: [
  //       '<Field Name="FirstTimeInstallation" ID="{BE0DAF97-33A2-4453-ACB2-FF99AF12C86A}" DisplayName="FirstTimeInstallation" Type="Boolean"/>',
  //     ],
  //     hasDefaultValues: true,
  //     defaultValues: [
  //       {
  //         Title: "RMSettingsData",
  //         DateFormat: "MM/DD/YYYY",
  //         ThemeMode: "theme",
  //         HideSPFxDefaultComponents: JSON.stringify({
  //           WebpartCustomCss: true,
  //           HideWebpartTitle: true,
  //           HideSideNavBar: true,
  //           HideTopCommandBar: true,
  //           HideTopSiteHeader: true,
  //           HideCommentsWrapper: true,
  //         }),
  //         Language: "English",
  //         FavIcon: true,
  //         FirstTimeInstallation: true,
  //       },
  //     ],
  //   },
  // ];

  // private async createColumns(listname: string, colLength: any): Promise<void> {
  //   let web = new Web(ContextService.GetUrl());
  //   return await web.lists.ensure(listname).then((ler: ListEnsureResult) => {
  //     const batch = web.createBatch();
  //     for (let i = 0; i < colLength.length; i++) {
  //       ler.list.fields
  //         .inBatch(batch)
  //         .createFieldAsXml(colLength[i])
  //         .catch((e: any) => {
  //           return e;
  //         });
  //     }
  //     return batch.execute();
  //   });
  // }

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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
