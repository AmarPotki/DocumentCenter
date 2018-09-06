import * as React from "react";
import * as ReactDom from "react-dom";
import { Version, UrlQueryParameterCollection } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from "@microsoft/sp-webpart-base";
import * as strings from "DocumentDetailWebPartStrings";
import { Main } from "./components/Main";
import { IMainProps } from "./components/Main/IMainProps";
import { IDocumentDetailWebPartProps } from "./IDocumentDetailWebPartProps";
import { IDropdownOption } from "office-ui-fabric-react";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import {
  IServiceFactory, ServiceFactory, MockDocumentService, MockListService,
  IListService, ListService, IDocumentService, DocumentService, ISearchService, SearchService
} from "../../services";
import { IPage, IList } from "../../domains";
import { IRefinementValue } from "../../domains/IRefinementValue";
import { PropertyPaneTaxonomyPicker } from "../../controls/taxonomyPicker/PropertyPaneTaxonomyPicker";
import { ITaxonomyValue } from "react-taxonomypicker";
import { forIn } from "@microsoft/sp-lodash-subset";

export default class DocumentDetailWebPart extends BaseClientSideWebPart<IDocumentDetailWebPartProps> {
  private searchService: ISearchService;
  private pagesOptions: IDropdownOption[];
  private listsOptions: IDropdownOption[];
  private documentService: IDocumentService;
  private listService: IListService;
  private serviceFactory: IServiceFactory;
  private managedPropertiesOptions: ITaxonomyValue[];
  /**
   * Renders the web part
   */

  public render(): void {

    this.searchService = new SearchService(this.context);
    this.documentService = new DocumentService(this.context);
    if (this.managedPropertiesOptions === undefined) {
      this.searchService.GetAllManagedProperties()
        .then((response: IRefinementValue[]) => {
          this.managedPropertiesOptions = response.map((item: IRefinementValue) => {
            return <ITaxonomyValue>{
              value: item.RefinementValue,
              label: item.RefinementName
            };
          });
        });
    }
    var queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
    let docPath: string = "";
    let resultPage: string = "";
    if (queryParameters.getValue("UniqueId")) {
      docPath = queryParameters.getValue("UniqueId");
    }
    if (queryParameters.getValue("Source")) {
      resultPage = decodeURIComponent(queryParameters.getValue("Source"));
    } else {
      resultPage = `${this.context.pageContext.web.absoluteUrl}/pages/${this.properties.searchResultPage}`;
    }
    // creating the React controls
    const element: React.ReactElement<IMainProps> = React.createElement(
      Main,
      {
        documentPath: docPath,
        noDocumentsMessage: this.properties.noDocumentMessage,
        showDescription: this.properties.showDescription,
        documentService: this.documentService,
        searchResultsPage: resultPage,
        shareBodyContent: this.properties.shareBodyContent,
        searchService: this.searchService,
        managedProperties: this.properties.managedProperties,
        managedPropertyForDescription: this.properties.managedPropertyForDescription,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,

        // createdDateLabel:this.properties.createdDateLabel,
        // categoryLabel:this.properties.categoryLabel,
        // issueDateLabel:this.properties.issueDateLabel
      }
    );
    // rendering the web part
    ReactDom.render(element, this.domElement);
  }

  /**
   * loads pages for 'searchResultPage' property
   * you can call your Promise methods here
   */
  // protected async onInit<T>(): Promise<void> {
  //   // initializing the services
  //   this.serviceFactory = new ServiceFactory();
  //   this.z = this.serviceFactory.GetService<IListService>(this.context, ListService, MockListService);
  //   this.documentService = this.serviceFactory.GetService<IDocumentService>(this.context, DocumentService, MockDocumentService);
  //   // get the pages to be loaded in "searchResultPage" property
  //   return new Promise<void>(resolve => {
  //     this.listService.getPages("Pages")
  //       .then((response) => {
  //         this.pagesOptions = response.map((item: IPage) => {
  //           return {
  //             key: item.Name,
  //             text: item.Name
  //           };
  //         });
  //         resolve();
  //       });
  //   });

  // }

  /**
   * web part version
   */
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  /**
   * check web part dependencies and validate the properties
   */
  private needsConfiguration(): boolean {
    return this.properties.searchResultPage === null ||
      this.properties.searchResultPage === undefined ||
      this.properties.searchResultPage.trim().length === 0;
  }

  /**
   * update properties value when it changed
   */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === "searchResultPage" && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }
  private _updateTaxonomyPicker = (name, value) => {
    if (value !== null && value !== undefined) {
      if (value.hasOwnProperty("length")) {
        this.properties.managedProperties = value.map((item) => item.label).join(",");
        this.properties.defaultValue = value;
      } else {
        this.properties.managedProperties = value.toString();
      }
      this.render();
    }
  }
  private _updateTaxonomyPickerDescription = (name, value) => {
    if (value !== null && value !== undefined) {
      this.properties.managedPropertyForDescription = value.value;
      this.properties.defaultValueForDescription = value;
    } else {
      this.properties.managedPropertyForDescription = "HitHighlightedSummary";
    }
    this.render();
  }
  protected onPropertyPaneConfigurationStart(): void {
    this.serviceFactory = new ServiceFactory();
    this.listService = this.serviceFactory.GetService<IListService>(this.context, ListService, MockListService);
    if (!this.listsOptions) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, "listPath");
      this.listService.getLists()
        .then((response) => {

          this.listsOptions = response.map((item: IList) => {
            return {
              key: item.ServerRelativeUrl,
              text: item.Title
            };
          });
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // this.render();
        });
    }
    if (!this.pagesOptions) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, "documentDetailPage");
      this.listService.getPages("Pages")
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
    }
    // if (this.properties.managedPropertiesOptions === undefined) {
    // this.context.statusRenderer.displayLoadingIndicator(this.domElement, "managedProperties");
    // this.searchService.GetAllManagedProperties()
    //   .then((response: IRefinementValue[]) => {
    //     // console.log(response);
    //     // this.managedPropertiesOptions = response.map((item: IRefinementValue) => {
    //     //   return <ITaxonomyValue>{
    //     //     value: item.RefinementValue,
    //     //     label: item.RefinementName
    //     //   };
    //     // });
    //     let arr = []
    //     this.managedPropertiesOptions = [];
    //     for (let item of response) {
    //       arr.push({ value: item.RefinementValue, label: item.RefinementName });
    //     }
    //     this.managedPropertiesOptions = arr;
    //     this.context.propertyPane.refresh();
    //     this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    //     // this.render();
    //   });
    // }

    // this.managedPropertiesOptions = [<ITaxonomyValue>{ value: "Ashkan", label: "Ashkan" },
    // <ITaxonomyValue>{ value: "Amar", label: "Amar" }];
  }
  /**
   * Initilize web part properties
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        // page 1
        {
          // header: {
          //   description: strings.PropertyPaneDescription
          // },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown("searchResultPage", {
                  label: strings.SearchResultPageLabel,
                  options: this.pagesOptions,
                  selectedKey: this.properties.searchResultPage
                }),
                PropertyPaneToggle("showDescription", {
                  label: strings.ShowDescriptionLabel
                }),
                PropertyPaneTextField("noDocumentMessage", {
                  label: strings.NoDocumentMessageLabel
                }),
                PropertyPaneTaxonomyPicker("managedProperties", {
                  // onRender: () => {},
                  key: "value",
                  name: "label",
                  displayName: "Select desire Properties",
                  multi: true,
                  termSetCountMaxSwapToAsync: 200,
                  defaultOptions: this.managedPropertiesOptions,
                  onPickerChange: this._updateTaxonomyPicker,
                  defaultValue: this.properties.defaultValue,
                }),
                PropertyPaneTaxonomyPicker("managedPropertyForDescription", {
                  key: "value",
                  name: "label",
                  displayName: strings.managedPropertyForDescriptionLabel,
                  multi: false,
                  termSetCountMaxSwapToAsync: 200,
                  defaultOptions: this.managedPropertiesOptions,
                  onPickerChange: this._updateTaxonomyPickerDescription,
                  defaultValue: this.properties.defaultValueForDescription,
                }),
                PropertyPaneTextField("shareBodyContent", {
                  label: "Email body message",
                  multiline: true
                })
              ]
            }
          ]
        },
        // page 2
        // {
        //   header: {
        //     description: "Managed Properties Configuration"
        //   },
        //   groups: [
        //     {
        //       groupName: "Managed Properties",
        //       groupFields: [

        //       ]
        //     }
        //   ]
        // }
      ]
    };
  }
}