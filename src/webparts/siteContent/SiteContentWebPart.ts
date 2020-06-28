import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "SiteContentWebPartStrings";
import SiteContent from "./components/SiteContent";
import { ISiteContentProps } from "./components/ISiteContentProps";

import HttpService from "./services/HttpService";

export interface ISiteContentWebPartProps {
  viewSiteContentBy: string;
  orderBy: string;
}

export default class SiteContentWebPart extends BaseClientSideWebPart<
  ISiteContentWebPartProps
> {
  protected async onInit(): Promise<void> {
    HttpService.Init(this.context.httpClient);
  }
  public render(): void {
    const element: React.ReactElement<ISiteContentProps> = React.createElement(
      SiteContent,
      {
        siteUrl: this.context.pageContext.web.absoluteUrl,
        orderBy: this.properties.orderBy,
        viewSiteContentBy: this.properties.viewSiteContentBy,
      },
      {
        baseTemplateId: "",
        title: "",
        url: "",
        itemCount: "",
        lastModifiedDate: "",
        createdDate: "",
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
  public getOrderBy(): IPropertyPaneDropdownOption[] {
    return [
      { key: "ModifiedDesc", text: "Modified Date Descending" },
      { key: "ModifiedAsc", text: "Modified Date Ascending" },
      { key: "CreatedDesc", text: "Created Date Descending" },
      { key: "CreatedAsc", text: "Created Date Ascending" },
      { key: "TitleDesc", text: "Title Descending" },
      { key: "TitleAsc", text: "Title Ascending" },
    ];
  }

  public getSiteContentOptions(): IPropertyPaneDropdownOption[] {
    return [
      { key: "lists", text: "Lists" },
      { key: "libraries", text: "Libraries" },
      { key: "all", text: "All" },
    ];
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                PropertyPaneDropdown("viewSiteContentBy", {
                  label: "View Site Content By",
                  options: this.getSiteContentOptions(),
                  selectedKey: "all",
                }),

                PropertyPaneDropdown("orderBy", {
                  label: "Order By",
                  selectedKey: "CreatedDesc",
                  options: this.getOrderBy(),
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
