import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import * as strings from 'RelatedDocumentsWebPartStrings';
import RelatedDocuments from './components/RelatedDocuments';
import { IRelatedDocumentsProps } from './components/IRelatedDocumentsProps';
import { IRelatedDocumentsWebPartProps } from './IRelatedDocumentsWebPartProps';
import { IDropdownOption } from 'office-ui-fabric-react';
import { IDocumentService, DocumentService, IServiceFactory, ServiceFactory, MockDocumentService, MockListService, IListService, ListService } from '../../services';
import { IList } from '../../domains';

export default class RelatedDocumentsWebPart extends BaseClientSideWebPart<IRelatedDocumentsWebPartProps> {
  private documentService: IDocumentService;
  private listService: IListService;
  private listsOptions: IDropdownOption[];
  private serviceFactory: IServiceFactory;

  constructor() {
    super();
    this.serviceFactory = new ServiceFactory();
    this.listService = this.serviceFactory.GetService<IListService>(this.context, ListService, MockListService);
    this.documentService = this.serviceFactory.GetService<IDocumentService>(this.context, DocumentService, MockDocumentService);
  }

  public render(): void {
    // get current document path from QueryString
    var queryParameters = new UrlQueryParameterCollection(window.location.href);
    let docPath: string = "";
    // set the docPath variable if the query parameter is not null or empty
    if (queryParameters.getValue("DocumentPath"))
      docPath = queryParameters.getValue("DocumentPath");

    const element: React.ReactElement<IRelatedDocumentsProps> = React.createElement(
      RelatedDocuments,
      {
        webPartProps: this.properties,
        documentService: this.documentService,
        documentPath: docPath
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listPath' && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'library');

    this.listService.getLists()
      .then((response) => {
        this.listsOptions = response.map((item: IList) => {
          return {
            key: item.Title,
            text: item.Title
          };
        });
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });

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
                PropertyPaneDropdown('library', {
                  label: strings.library,
                  options: this.listsOptions,
                  selectedKey: this.properties.library
                }),
                PropertyPaneTextField('top', {
                  label: strings.top
                }),
                PropertyPaneTextField('orderByField', {
                  label: strings.order
                }),
                PropertyPaneToggle('ascending', {
                  label: strings.ascending
                }),
                PropertyPaneTextField('noDocumentMessage', {
                  label: strings.noDocsMessage
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
