import * as React from "react";
import styles from "./RelatedDocuments.module.scss";
import { IRelatedDocumentsProps } from "./IRelatedDocumentsProps";
import { IRelatedDocumentsState } from "./IRelatedDocumentsState";
// import { escape } from "@microsoft/sp-lodash-subset";
import { SearchQueryBuilder } from "sp-pnp-js";
import { IDocumentItem } from "../../../domains/IDocumentItem";
export default class RelatedDocuments extends React.Component<IRelatedDocumentsProps, IRelatedDocumentsState> {
  private timerID: any;
  constructor() {
    super();
    this.state = { documents: [], containerWidth: 0 };
  }
  public componentDidMount() {
    this.timerID = setInterval(
      () => this.watchSize(),
      1000
    );
  }
  protected watchSize(): void {
    const width = document.getElementById(styles.relatedDocuments).clientWidth;
    this.setState({ containerWidth: width });
  }
  public componentWillUnmount() {
    clearInterval(this.timerID);
  }
  public render(): React.ReactElement<IRelatedDocumentsProps> {
    this.getRelatedDocuments();
    return (
      <div className={styles.relatedDocuments} id={styles.relatedDocuments}>
        <div className="documents-block">
          <ul className="documents-list recent-added">
            {
              this.state.documents.length < 1 &&
              this.props.webPartProps.noDocumentMessage
            }
            {
              this.state.documents
            }
          </ul>
        </div>
      </div>
    );
  }

  private async getRelatedDocuments(): Promise<void> {
    let documentsElement: any[] = [];
    let documents: IDocumentItem[] = await this.props.documentService.getRelatedDocumentsByPath(
      this.props.webPartProps.library, this.props.documentPath,
      this.props.webPartProps.top, this.props.webPartProps.orderByField,
      this.props.webPartProps.ascending);

    for (let i: number = 0; i < documents.length; i++) {
      let document: IDocumentItem = documents[i];
      documentsElement.push(
        <li>
          {/* <a href={document.File.ServerRelativeUrl}>
            <div className={styles.container}><img src={document.Icon} />
              <span>{document.Title}</span>
            </div>
          </a> */}
        </li>);
    }

    this.setState({
      documents: documentsElement
    });

    // return new Promise((resolve, reject) => {
    //   resolve();
    // });
  }
}
