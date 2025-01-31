import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import * as strings from "EmpApiWebPartStrings";
import EmpApi from "./components/EmpApi";
import { IEmpApiProps } from "./components/IEmpApiProps";

export interface IEmpApiWebPartProps {
  description: string;
}

export default class EmpApiWebPart extends BaseClientSideWebPart<IEmpApiWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    // Pass necessary props to the EmpApi component
    const element: React.ReactElement<IEmpApiProps> = React.createElement(EmpApi, {
      description: this.properties.description || "No description provided", // Description field
      isDarkTheme: this._isDarkTheme, // Pass dark theme state
      environmentMessage: this._environmentMessage, // Pass environment message
      hasTeamsContext: !!this.context.sdks?.microsoftTeams, // Check if it's in Teams context
      userDisplayName: this.context.pageContext.user.displayName || "Guest", // Display user name
      context: this.context, // Pass SharePoint context
    });

    ReactDom.render(element, this.domElement); // Render EmpApi component
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = await this._getEnvironmentMessage();
  }

  // Get environment message based on the context
  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks?.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((context) => {
        switch (context.app.host.name) {
          case "Office":
            return this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentOffice
              : strings.AppOfficeEnvironment;
          case "Outlook":
            return this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentOutlook
              : strings.AppOutlookEnvironment;
          case "Teams":
          case "TeamsModern":
            return this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentTeams
              : strings.AppTeamsTabEnvironment;
          default:
            return strings.UnknownEnvironment;
        }
      });
    }
    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  // Handle theme change (dark/light theme)
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement); // Clean up when disposed
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
