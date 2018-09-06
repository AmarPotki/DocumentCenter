import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
//  import { initializeIcons } from '@uifabric/icons';
//  initializeIcons();
import * as strings from 'ResultsFilterWebPartStrings';
import ResultsFilter from './components/ResultsFilter/ResultsFilter';
import { IResultsFilterProps } from './components/ResultsFilter/IResultsFilterProps';
import { IResultsFilterWebPartProps } from './IResultsFilterWebPartProps';
import { PropertyFieldTermSetPicker } from '../../controls/termSetPicker';
import ISPTermStorePickerService from '../../services/SPTermStorePickerService';
import SPTermStorePickerService from '../../services/SPTermStorePickerService';
export default class ResultsFilterWebPart extends BaseClientSideWebPart<IResultsFilterWebPartProps> {
  private termStoreService: ISPTermStorePickerService;
  public render(): void {
    this.termStoreService = new SPTermStorePickerService(this.context);
    const element: React.ReactElement<IResultsFilterProps> = React.createElement(
      ResultsFilter,
      {
        terms: this.properties.terms || [],
        termStorePickerService: this.termStoreService,
        title: this.properties.title
      }
    );

    ReactDom.render(element, this.domElement);
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
              //  groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleLable,
                }),
                PropertyFieldTermSetPicker('terms', {
                  label: 'Select terms',
                  panelTitle: 'Select terms',
                  initialValues: this.properties.terms,
                  allowMultipleSelections: true,
                  excludeSystemGroup: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'termSetsPickerFieldId',

                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
