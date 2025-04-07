import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
// import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "IfcDocumentViewerWebPartStrings";
import DocumentLibraryViewer from "./components/IfcDocumentViewer";
import { IDocumentLibraryViewerProps } from "./components/IIfcDocumentViewerProps";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IDocumentLibraryViewerWebPartProps {
  title: string;
  description: string;
  libraryName: string;
}

export default class DocumentLibraryViewerWebPart extends BaseClientSideWebPart<IDocumentLibraryViewerWebPartProps> {
  private _libraries: { key: string; text: string }[] = [];
  // private _isDarkTheme: boolean = false;
  // private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<IDocumentLibraryViewerProps> =
      React.createElement(DocumentLibraryViewer, {
        title: this.properties.title,
        description: this.properties.description,
        context: this.context,
        libraryName: this.properties.libraryName,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
      });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await this._getLibraries();
    return super.onInit();
  }

  private async _getLibraries(): Promise<void> {
    const response = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 101`,
      SPHttpClient.configurations.v1
    );

    if (response.ok) {
      const data = await response.json();
      this._libraries = data.value.map((list: { Title: unknown }) => ({
        key: list.Title,
        text: list.Title,
      }));
    } else {
      console.error("Error fetching libraries:", response.statusText);
    }
  }

  // protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
  //   if (!currentTheme) {
  //     return;
  //   }

  //   this._isDarkTheme = !!currentTheme.isInverted;
  //   const { semanticColors } = currentTheme;

  //   if (semanticColors) {
  //     this.domElement.style.setProperty(
  //       "--bodyText",
  //       semanticColors.bodyText || null
  //     );
  //     this.domElement.style.setProperty("--link", semanticColors.link || null);
  //     this.domElement.style.setProperty(
  //       "--linkHovered",
  //       semanticColors.linkHovered || null
  //     );
  //   }
  // }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("title", {
                  label: "Web Part Title",
                }),
                PropertyPaneTextField("description", {
                  label: "Web Part Description",
                  multiline: true,
                }),
                PropertyPaneDropdown("libraryName", {
                  label: "Select Document Library",
                  options: this._libraries,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
