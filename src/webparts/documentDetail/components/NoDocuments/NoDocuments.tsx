import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { INoDocumentsProps } from './INoDocumentsProps';

export class NoDocuments extends React.Component<INoDocumentsProps, {}> {
  /**
   * Renders the control
   */
  public render(): JSX.Element {
    return (
      <div>
          {this.props.noDocumentsMessage}
      </div>
    );
  }
}