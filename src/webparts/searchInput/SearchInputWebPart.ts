import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SearchInputWebPartStrings';
import SearchInput from './components/SearchInput';
import { ISearchInputProps } from './components/ISearchInputProps';

export interface ISearchInputWebPartProps {
  placeholder: string;
  title:string;
}

export default class SearchInputWebPart extends BaseClientSideWebPart<ISearchInputWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISearchInputProps > = React.createElement(
      SearchInput,
      {
        placeholder: this.properties.placeholder,
        keyword:strings.Keyword,
        title:this.properties.title
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
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleLable,
                }),
                PropertyPaneTextField('placeholder', {
                  label: strings.PlaceholderLable,
                  value:strings.Placeholder
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
