import * as React from 'react';
import styles from './ResultsFilter.module.scss';
import { IResultsFilterProps } from './IResultsFilterProps';
import TermsSetDropDown from '../termSetDropDown/TermsSetDropDown';
import { ICheckedTermSets, ICheckedTermSet } from '../../../../controls/termSetPicker';
// export interface IResultsFilterState {
//   containerWidth: number;
// }
export default class ResultsFilter extends React.Component<IResultsFilterProps, any> {
  private timerID: any;
  /**
   *
   */
  constructor(props: IResultsFilterProps) {
    super(props);
    this.state = { containerWidth: 0, };
  }
  public componentDidMount() {
    this.timerID = setInterval(
      () => this.watchSize(),
      1000
    );
  }
  public componentWillUnmount() {
    clearInterval(this.timerID);
  }
  protected watchSize(): void {
    const width = document.getElementById(styles.container).clientWidth;
    this.setState({ containerWidth: width });
  }

  public render(): React.ReactElement<IResultsFilterProps> {
    return (
      <div className={styles.resultsFilter} >
        {/* id uses in watchSize methods */}
        <div className={styles.container} id={styles.container}>
          <div className="ms-Grid">
            <div className={`ms-Grid-row ${styles.row}`}>
              <div className={`ms-Grid-col ms-lg4 ms-md6 ms-sm12 ${styles.column}`}>
                <div className={styles.applyFilter}>
                  {this.props.title}
                </div>
              </div>
            </div>

            <div className={styles.row}>
              {this.props.terms.map((term: ICheckedTermSet, key: number) => {
                if (this.state.containerWidth > 420) {
                  return (<div className={`ms-Grid-col ms-lg4 ms-md6 ms-sm12 ${styles.column}`}>
                    <TermsSetDropDown {...this.props} termSet={term} key={key} />
                  </div>);
                } else {
                  return (<div className={`ms-Grid-col ms-lg12 ms-md12 ms-sm12 ${styles.column}`}>
                    <TermsSetDropDown {...this.props} termSet={term} key={key} />
                  </div>);
                }
              }
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }

}
