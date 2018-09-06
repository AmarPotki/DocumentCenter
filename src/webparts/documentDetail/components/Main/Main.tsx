import * as React from 'react';
import { IMainProps } from './IMainProps';
import { NoDocuments } from '../NoDocuments';
import { DocumentDetail } from '../DocumentDetail';
import styles from '../DocumentDetail.module.scss';
const backIcon: any = require("../../../../assets/arrow-prev-large.png");
export class Main extends React.Component<IMainProps, {}> {

   /**
   * constructor
   * @param props: control properties
   */
    constructor(props: IMainProps) {
        super(props);
    }
    /**
   * Renders the control
   */
    public render(): JSX.Element {
        return (
            <div className={styles.documentDetail}>
                <div className={styles.backToResultPage}>
                    <a href={this.props.searchResultsPage}>
                        <img src={backIcon} alt="Back to result page" />
                        <p>Back to search results</p>
                    </a>
                </div>
                {/* <div className={styles.documentDetailContainer}> */}
                <div>
                    {
                        // Render NoDocuments component if no document path defined in the url
                        this.props.documentPath.length <1 &&
                        <NoDocuments noDocumentsMessage={this.props.noDocumentsMessage} />
                    }
                    {
                        // Render DocumentDetail if document path defined properly
                        this.props.documentPath.length > 1 &&
                        <DocumentDetail {...this.props} />
                    }
                </div>
            </div>
        );
    }
}