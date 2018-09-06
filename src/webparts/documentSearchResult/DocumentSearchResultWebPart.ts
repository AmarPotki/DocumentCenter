import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";
import * as strings from "DocumentSearchResultWebPartStrings";
import { Main } from "./components/Main";
import { IMainProps } from "./components/Main/IMainProps";
import { IDocumentSearchResultWebPartProps } from "./IDocumentSearchResultWebPartProps";
import {
  IDocumentService, ISearchService, SearchService, IListService,
  ServiceFactory, IServiceFactory, ListService, MockListService
} from "../../services";
import { ITaxonomyHelper } from "../../common/ITaxonomyHelper";
import { IPage, IList } from "../../domains";
import { TaxonomyHelper } from "../../common/TaxonomyHelper";
import { DocumentService } from "../../services/DocumentService";
import { IRefinementValue } from "../../domains/IRefinementValue";
import { PropertyPaneTaxonomyPicker } from "../../controls/taxonomyPicker/PropertyPaneTaxonomyPicker";
import { ITaxonomyValue } from "react-taxonomypicker";
import { ISearchSource } from "../../domains/ISearchSource";

export default class DocumentSearchResultWebPart extends BaseClientSideWebPart<IDocumentSearchResultWebPartProps> {
  private searchService: ISearchService;
  private taxonomyHelper: ITaxonomyHelper;
  private listService: IListService;
  private pagesOptions: IDropdownOption[];
  private listsOptions: IDropdownOption[];
  private managedPropertiesOptions: ITaxonomyValue[];
  private serviceFactory: IServiceFactory;
  private documentService: IDocumentService;
  constructor() {
    super();
  }

  public render(): void {
    this.searchService = new SearchService(this.context);
    this.taxonomyHelper = new TaxonomyHelper(this.context);
    this.documentService = new DocumentService(this.context);

    // if (!this.managedPropertiesOptions) {
    //   this.context.statusRenderer.displayLoadingIndicator(this.domElement, "managedProperties");
    //   this.searchService.GetAllManagedProperties()
    //     .then((response) => {
    //       this.managedPropertiesOptions = response.map((item: IRefinementValue) => {
    //         return <ITaxonomyValue>{
    //           value: item.RefinementValue,
    //           label: item.RefinementName
    //         };
    //       });
    //       this.context.propertyPane.refresh();
    //       this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    //       // this.render();
    //     }).catch(console.log);
    // }

    const element: React.ReactElement<IMainProps> = React.createElement(
      Main,
      {
        context: this.context,
        searchService: this.searchService,
        taxonomyHelper: this.taxonomyHelper,
        webPartProperties: this.properties,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        documentService: this.documentService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.serviceFactory = new ServiceFactory();
    this.listService = this.serviceFactory.GetService<IListService>(this.context, ListService, MockListService);
    if (!this.listsOptions) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, "listPath");
      this.searchService.GetSearchSources()
        .then((response) => {
          this.listsOptions = response.map((item: ISearchSource) => {
            return <IDropdownOption>{
              key: item.Id,
              text: item.Name
            };
          });
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        });
    }
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
          this.render();
        });
    }
    if (!this.managedPropertiesOptions) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, "managedProperties");
      this.searchService.GetAllManagedProperties()
        .then((response) => {
          this.managedPropertiesOptions = response.map((item: IRefinementValue) => {
            return <ITaxonomyValue>{
              value: item.RefinementValue,
              label: item.RefinementName
            };
          });
          console.log("Manages Properties Refresh Property Pane");
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        }).catch(console.log);
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === "listPath" && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }

    this.render();
  }

  private _updateTaxonomyPicker = (name, value) => {
    if (value !== null && value !== undefined) {
      if (value.hasOwnProperty("length")) {
        this.properties.managedProperties = value.map((item) => item.label).join(",");
        this.properties.managedPropertiesDefaultValue = value;
      } else {
        this.properties.managedProperties = value.toString();
      }
      this.render();
    }
  }

  private _updateListPicker = (name, value) => {
    if (value !== null && value !== undefined) {
      if (value.hasOwnProperty("length")) {
        this.properties.listPath = value.map((item) => item.value).join(",");
        this.properties.listsPathDefaultValue = value;
      } else {
        this.properties.listPath = value.toString();
      }
      this.render();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        // page 1
        {
          // header: {
          //   description: strings.PageOneDescription
          // },
          groups: [
            {
              groupName: "Basic Settings", // strings.PageOneLookAndFeelName,
              groupFields: [
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("searchText", {
                //   label: "Search box text"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("searchButtonText", {
                //   label: "Search button text"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneToggle("showCategoryFilter", {
                //   key: "showCategoryFilter",
                //   label: "Show category filter?"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneToggle("showFunctionFilter", {
                //   key: "showFunctionFilter",
                //   label: "Show function filter?"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneToggle("showLocationFilter", {
                //   key: "showLocationFilter",
                //   label: "Show location filter?"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneToggle("showFileTypeFilter", {
                //   key: "showFileTypeFilter",
                //   label: "Show file type filter?"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneDropdown("listPath", {
                //   label: strings.ListPathLabel,
                //   options: this.listsOptions,
                //   selectedKey: this.properties.listPath
                // }),
                PropertyPaneTextField("listPath", {
                  label: "Search resource"
                }),
                PropertyPaneTaxonomyPicker("managedProperties", {
                  // tslint:disable-next-line:no-empty
                  // onRender: () => {},
                  key: "value",
                  name: "label",
                  displayName: "Managed properties",
                  multi: true,
                  termSetCountMaxSwapToAsync: 100,
                  defaultOptions: this.managedPropertiesOptions,
                  onPickerChange: this._updateTaxonomyPicker,
                  defaultValue: this.properties.managedPropertiesDefaultValue
                }),
                PropertyPaneDropdown("sortDocumentBy", {
                  label: "Default sorting",
                  options: [
                    { key: "LastModifiedTime", text: "Modified" },
                    { key: "Created", text: "Created" },
                    { key: "RefinableString06", text: "Title" }
                  ]
                }),
                PropertyPaneToggle("showShareButton", {
                  key: "showShareButton",
                  label: "Show share button?"
                }),
                PropertyPaneToggle("showDescription", {
                  key: "showDescription",
                  label: "Show description?"
                }),
                PropertyPaneToggle("enableSearchOrdering", {
                  key: "enableSearchOrdering",
                  label: "Enable search ordering?"
                }),
                PropertyPaneToggle("useDocumentDetailPage", {
                  key: "useDocumentDetailPage",
                  label: "Use document detail page?"
                }),
                PropertyPaneDropdown("documentDetailPage", {
                  label: strings.DocumentDetailPageLabel,
                  options: this.pagesOptions,
                  selectedKey: this.properties.documentDetailPage,
                  disabled: !this.properties.useDocumentDetailPage
                }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("functionFilterLabel", {
                //   label: "Function filter label"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("categoryFilterLabel", {
                //   label: "Category filter label"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("locationFilterLabel", {
                //   label: "Location filter label"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("fileTypeFilterLabel", {
                //   label: "FileType filter label"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("createdDateLabel", {
                //   label: "Created date label"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("issueDateLabel", {
                //   label: "Issue date Label"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("categoryLabel", {
                //   label: "Document category label"
                // }),
                // tslint:disable-next-line:comment-format
                // PropertyPaneTextField("shareBodyContent", {
                //   label: "Email body message",
                //   multiline: true
                // })
              ]
            }
          ]
        },
        // page 2
        // {
        //   header: {
        //     description: strings.PageTwoDescription
        //   },
        //   groups: [
        //     {
        //       groupName: strings.PageTwoGroupName,
        //       groupFields: [
        //         // tslint:disable-next-line:comment-format
        //         // PropertyPaneDropdown("listPath", {
        //         //   label: strings.ListPathLabel,
        //         //   options: this.listsOptions,
        //         //   selectedKey: this.properties.listPath
        //         // }),
        // tslint:disable-next-line:comment-format
        //         PropertyPaneTaxonomyPicker("listPath", {
        //           // tslint:disable-next-line:no-empty
        //           // onRender: () => {},
        //           key: "value",
        //           name: "label",
        //           displayName: strings.ListPathLabel,
        //           multi: true,
        //           termSetCountMaxSwapToAsync: 100,
        //           defaultOptions: this.listsOptions,
        //           onPickerChange: this._updateListPicker,
        //           defaultValue: this.properties.listsPathDefaultValue
        //         }),
        // tslint:disable-next-line:comment-format
        //         PropertyPaneDropdown("documentDetailPage", {
        //           label: strings.DocumentDetailPageLabel,
        //           options: this.pagesOptions,
        //           selectedKey: this.properties.documentDetailPage
        //         }),
        // tslint:disable-next-line:comment-format
        //         PropertyPaneDropdown("sortDocumentBy", {
        //           label: "Sort document by",
        //           options: [
        //             { key: "LastModifiedTime", text: "Modified" },
        //             { key: "Created", text: "Created" },
        //             { key: "RefinableString06", text: "Title" }
        //           ]
        //         }) // ,
        //         // tslint:disable-next-line:comment-format
        //         // PropertyPaneTextField("defaultQueryText", {
        //         //   label: "Query text",
        //         //   multiline: true
        //         // })
        //       ]
        //     }
        //   ]
        // },
        // page 3
        // {
        //   header: {
        //     description: strings.PageThreeDescription
        //   },
        //   groups: [
        //     {
        //       groupName: strings.PageThreeGroupName,
        //       groupFields: [
        // tslint:disable-next-line:comment-format
        //         PropertyPaneTaxonomyPicker("managedProperties", {
        //           // tslint:disable-next-line:no-empty
        //           // onRender: () => {},
        //           key: "value",
        //           name: "label",
        //           displayName: "Select desire Properties",
        //           multi: true,
        //           termSetCountMaxSwapToAsync: 100,
        //           defaultOptions: this.managedPropertiesOptions,
        //           onPickerChange: this._updateTaxonomyPicker,
        //           defaultValue: this.properties.managedPropertiesDefaultValue
        //         })
        //       ]
        //     }
        //   ]
        // }
      ]
    };
  }
}
