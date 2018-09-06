import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown, PropertyPaneLabel
} from "@microsoft/sp-webpart-base";

import * as strings from "RecentDocumentsWebPartStrings";
import RecentDocuments from "./components/RecentDocuments";
import { IRecentDocumentsProps } from "./components/IRecentDocumentsProps";
import { ISearchService, SearchService, IListService, ListService, IServiceFactory, MockListService, ServiceFactory, GraphApiService } from "../../services";
import { IRecentDocumentsWebPartProps } from "./IRecentDocumentsWebPartProps";
import { PropertyPaneTaxonomyPicker } from "../../controls/taxonomyPicker/PropertyPaneTaxonomyPicker";
import { ITaxonomyValue } from "react-taxonomypicker";
import { IList } from "../../domains/IList";
import { IDropdownOption, IDropdown } from "office-ui-fabric-react";
import { IPage } from "../../domains/IPage";

export default class RecentDocumentsWebPart extends BaseClientSideWebPart<IRecentDocumentsWebPartProps> {
  private searchService: ISearchService;
  private listService: IListService;
  private serviceFactory: IServiceFactory;
  private pagesOptions: IDropdownOption[];
  private modeOption: IDropdownOption[];

  public render(): void {
    this.listService = new ListService(this.context);
    const element: React.ReactElement<IRecentDocumentsProps> = React.createElement(
      RecentDocuments,
      {
        recentDocumentCount: this.properties.recentDocumentCount,
        searchService: new SearchService(this.context),
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        webPartProperties: this.properties,
        graphService: new GraphApiService(this.context),
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
  protected onPropertyPaneConfigurationStart(): void {
    if (!this.pagesOptions) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, "documentDetailPage");

      this.listService.getPages("Site Pages")
        .then((response) => {
          this.pagesOptions = response.map((item: IPage) => {
            return {
              key: item.Name,
              text: item.Name
            };
          });

          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // this.render();
        });
      if (!this.modeOption) {
        this.context.statusRenderer.displayLoadingIndicator(this.domElement, "mode");
        this.modeOption = [];
        this.modeOption.push({ key: "Added", text: "Recently Added" });
        this.modeOption.push({ key: "Viewed", text: "Recently Viewed" });
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        // this.render();
      }
    }
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          // header: {
          //   description: strings.PropertyPaneDescription
          // },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleLable,
                }),
                PropertyPaneTextField("recentDocumentCount", {
                  label: strings.RecentDocumentCountFieldLabel,
                }),
                PropertyPaneDropdown("documentDetailPage", {
                  label: "Document Details Page",
                  options: this.pagesOptions,
                  selectedKey: this.properties.documentDetailPage
                }),
                PropertyPaneTextField("sourceName", {
                  label: strings.SourceNameLabel,
                }),
                PropertyPaneDropdown("mode", {
                  label: strings.ModeLable,
                  options: this.modeOption,
                  selectedKey: this.properties.mode,
                  
                }),
                PropertyPaneLabel("notification", {
                  text: this.properties.mode === "Viewed" ? "Mode not yet implemented" : "",
                  required: false
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
