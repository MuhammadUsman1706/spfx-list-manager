import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "HelloWorldWebPartStrings";
import { IHelloWorldProps } from "./components/IHelloWorldProps";
import Wrapper from "./components/HelloWorld";
import { ClientMode } from "./components/ClientMode";
import { AadHttpClient } from "@microsoft/sp-http";

export interface IHelloWorldWebPartProps {
  description: string;
  clientMode: ClientMode;
  sites: any[];
  ctx: any;
  selectedSite: any;
  selectedList: any;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  protected sites: any = [];
  protected lists: any = [];

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      Wrapper,
      {
        selectedSite: this.properties.selectedSite,
        selectedList: this.properties.selectedList,
        clientMode: this.properties.clientMode,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        ctx: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _searchWithAad = async (
    link: string,
    saveParam: string,
    filterFunction?: Function,
    method: "get" | "post" = "get",
    options?: any
  ): Promise<any> => {
    try {
      const client: AadHttpClient =
        await this.context.aadHttpClientFactory.getClient(
          "https://graph.microsoft.com"
        );

      const response = await client[method](
        link,
        AadHttpClient.configurations.v1,
        options
      );
      const json = await response.json();

      let filteredData: Array<any> = [];

      console.log(json);

      if (filterFunction) {
        filteredData = filterFunction(json.value);
        console.log(filteredData);
      } else {
        filteredData = json.value;
      }

      // if (saveParam) {
      //   this.context.dispatch({ type: saveParam, payload: filteredData });
      // }

      return filteredData;
    } catch (error) {
      console.error(error);
    }
  };

  protected async onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error("Unknown host");
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): Promise<any> {
    if (propertyPath === "selectedSite") {
      console.log(newValue.id);

      this.lists = await this._searchWithAad(
        `https://graph.microsoft.com/v1.0/sites/${newValue.id}/lists`,
        "SET_LISTS",
        (lists: any) =>
          lists.map((list: any) => ({
            key: list,
            text: list.displayName,
          }))
      );

      this.context.propertyPane.refresh();
    }

    if (propertyPath === "selectedList") {
      console.log("List changed!");
    }
  }

  protected async onPropertyPaneConfigurationStart(): Promise<any> {
    this.sites = await this._searchWithAad(
      "https://graph.microsoft.com/v1.0/sites?search=",
      "SET_SITES",
      (sites: any) =>
        sites.map((site: any) => ({
          key: site,
          text: site.displayName,
        }))
    );

    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneChoiceGroup("clientMode", {
                  label: strings.ClientModeLabel,
                  options: [
                    { key: ClientMode.aad, text: "AadHttpClient" },
                    { key: ClientMode.graph, text: "MSGraphClient" },
                  ],
                }),
                PropertyPaneDropdown("selectedSite", {
                  label: "Select a Site",
                  options: this.sites,
                  selectedKey: null,
                }),
                PropertyPaneDropdown("selectedList", {
                  label: "Select a List",
                  options: this.lists,
                  selectedKey: null,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

// "initialPage": "https://arkitektz.sharepoint.com/_layouts/15/sharepoint.aspx"
