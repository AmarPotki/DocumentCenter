import * as React from 'react';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { ITermProps, ITermState } from './IPropertyFieldTermSetPickerHost';

import styles from './PropertyFieldTermSetPickerHost.module.scss';
import { Label } from 'office-ui-fabric-react/lib/Label';


/**
 * Term component
 * Renders a selectable term
 */
export default class Term extends React.Component<ITermProps, ITermState> {
  constructor(props: ITermProps) {
    super(props);

    // Check if current term is selected
    let active = this.props.activeNodes.filter(item => item.id === this.props.term.Id);

    this.state = {
      selected: active.length > 0
    };

    
  }

  /**
   * Lifecycle event hook when component retrieves new properties
   * @param nextProps
   * @param nextContext
   */
  public componentWillReceiveProps?(nextProps: ITermProps, nextContext: any): void {
    // If multi-selection is turned off, only a single term can be selected
    if (!this.props.multiSelection) {
      let active = nextProps.activeNodes.filter(item => item.id === this.props.term.Id);
      this.state = {
        selected: active.length > 0
      };
    }
  }


  public render(): JSX.Element {
    const styleProps: React.CSSProperties = {
      marginLeft: `${(this.props.term.PathDepth * 30)}px`
    };

    return (
      <div className={`${styles.listItem} ${styles.term}`} style={styleProps}>
        <Label>{this.props.term.Name}</Label>
      </div>
    );
  }
}
