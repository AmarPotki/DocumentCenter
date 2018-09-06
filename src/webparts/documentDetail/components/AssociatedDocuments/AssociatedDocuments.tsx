import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IAssociatedDocumentsProps } from './IAssociatedDocumentsProps';
import { IAssociatedDocumentsState } from './IAssociatedDocumentsState';
import styles from '../DocumentDetail.module.scss';
import { IAssociatedDocument, ICCDocument } from '../../../../domains';
import { Log } from '@microsoft/sp-core-library';
import { IUtilities, Utilities } from "../../../../common";
import { SearchQuery, SearchResults, SearchQueryBuilder, Dictionary } from 'sp-pnp-js';
export class AssociatedDocuments extends React.Component<IAssociatedDocumentsProps, IAssociatedDocumentsState> {
  private utilities: IUtilities;
  /**
   * constructor
   * @param props: control properties
   */
  constructor(props: IAssociatedDocumentsProps) {
    super(props);
    this.utilities = new Utilities();
    // bind default values to current state
    this.state = {
      loading: true,
      error: undefined,
      associatedDocuments: undefined,
      documentsCount: 0,
      panelManagedProperties: [],
      results: null,
      content: '',
      displayAuthor: ''
    };
    // get associated documents
    this.getAssociatedDocuments();
  }

  /**
   * Renders the control
  */
  public render(): JSX.Element {
    // log to console
    Log.verbose("SpfxDocumentDetail", "Invoking render");
    return (
      this.state.documentsCount > 0 &&
      <div>
        <div className={styles.associatedDocsContainer}>

          Associated document(s) | {this.state.documentsCount} file(s) available
        </div>
        {this.state.associatedDocuments}
      </div>
    );
  }

  /**
   * Loads assicated documents to the web part
  */
  private async getAssociatedDocuments(): Promise<void> {
    try {
      // get ServerRelativeUrl of the primary document
      let url: string = this.props.primaryDocumentFolder;
      // call getAssociatedDocuments api to get the associated documents
      let associatedDocuments: IAssociatedDocument[] = await this.props.documentService.getAssociatedDocuments(url, this.props.webUrl);
      let docs: any[] = [];
      if (associatedDocuments.length > 0) {
        // associatedDocuments.forEach(async (element) => {
        for (let element of associatedDocuments) {
          let doc = await this.getDocument(element.UniqueId);
          docs.push(
            <div className={styles.hit}>
              <div className={styles.hitImage}>
                <img src={element.Icon} alt="icon" />
              </div>
              <div className={styles.hitContent}>
                <a className={styles.hitPrice}  ><img src={this.utilities.getIcon("Bookmark", false)} alt="Bookmark" /></a>
                <a className={styles.hitPrice} href={this.getMailTo(element.Path, element.Title)} ><img src={this.utilities.getIcon("Share", false)} alt="Share" /></a>
                <a className={styles.hitPrice}  ><img src={this.utilities.getIcon("Translate", false)} alt="Translate" /></a>
                <p className={styles.hitName}>
                  <a href={element.OnlinePath}>
                    {element.Title}
                  </a></p>
                <ul className={styles.documentMetaContainer}>
                  <li>
                    <strong>File Type: </strong> {element.FileType}</li>
                  <li>
                    <strong>Created: </strong> {element.Created}</li>
                  <li><strong>Modified: </strong> {element.TimeLastModified}</li>
                </ul>
                <p className={styles.hitDescription}>
                  {this.utilities.stripHTMLFromText(doc[0])}
                </p>
                <ul className={styles.documentMetaContainer}>
                  <li>
                    <img className={styles.user_icon}
                      src={doc[1]} />
                    <strong className={styles.created_user}>Author: </strong> {doc[2]}</li>
                  {(() => {
                    if (this.state.panelManagedProperties.length == 0) {
                      return;
                    }
                    let managedProperties: string[] = [];
                    managedProperties = this.props.managedProperties.split(",");
                    return (managedProperties.map((item: string) => {
                      {
                        if (this.state.results !== null) {
                          if (doc[3].PrimarySearchResults[0][item] !== "" && item !== "") {
                            return (<li><strong>{item}: </strong>
                              {doc[3].PrimarySearchResults[0][item]}   </li>);
                          }
                        }
                      }
                    }));
                  })()}
                </ul>
              </div>
            </div>
          );
        }
      }
      // set state to be included by document information
      this.setState({
        loading: false,
        error: "",
        associatedDocuments: docs,
        documentsCount: associatedDocuments.length
      });
    }
    catch (e) {
      // set state to returns error if an error occurred
      this.setState({
        error: "There are no available associated documents to display."
      });
      // log the error to console
      Log.error("SpfxDocumentDetail", e.message);
    }

  }
  private getMailTo(path: string, Title: string): string {
    let bodyContent: string = this.props.shareBodyContent;
    bodyContent = bodyContent.replace("##URL##", encodeURIComponent(path));
    bodyContent = bodyContent.replace(/(\r\n|\n|\r)/gm, "%0D%0A");
    let mailTo: string = `mailto:?subject=${encodeURIComponent(Title)}&body=${bodyContent}`;
    return mailTo;
  }
  private async getDocument(uniqueId: string): Promise<[string,string, string, SearchResults]> {
    try {
      const appSearchSettings: SearchQuery = {
        TrimDuplicates: false,
        Querytext: "*"
      };

      let myproperties: Array<string> = ["Title", "Size", "Filename",
        "FileType", "Created", "IssueDate", "DisplayAuthor", "AuthorOWSUSER", "ModifiedOWSDATE",
        "FileExtension", "Path", "ServerRedirectedEmbedURL", "HitHighlightedSummary", "SPWebUrl"];
      let managedProperties: string[] = [];
      if (this.props.managedProperties !== undefined) {
        managedProperties = this.props.managedProperties.split(",");
        managedProperties.forEach((key: string) => {
          if (!myproperties.some(x => x === key))
            myproperties.push(key);
        });
      }

      if (!myproperties.some(x => x === this.props.managedPropertyForDescription)) {
        myproperties.push(this.props.managedPropertyForDescription);
      }
      let query: SearchQueryBuilder = SearchQueryBuilder
        .create(`UniqueId: ${uniqueId}`, appSearchSettings)
        .rowLimit(1)
        .startRow(0)
        .selectProperties(...myproperties);
      // call

      // set state to be included by document information
      let results: SearchResults = await this.props.searchService.GetSearchResult(query);
      let tenantUrl: string = window.location.protocol + "//" + window.location.host;
      let documentPath = results.PrimarySearchResults[0]["Path"].replace(tenantUrl, "");

      let content = results.PrimarySearchResults[0][this.props.managedPropertyForDescription];
      let displayAuthor = results.PrimarySearchResults[0]["DisplayAuthor"];
      // doc.fileType = results.PrimarySearchResults[0]["FileType"];
      // doc.modified = results.PrimarySearchResults[0]["ModifiedOWSDATE"];
       let userIcon = `${this.props.webAbsoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${results.PrimarySearchResults[0]["AuthorOWSUSER"].split("|")[0].trim()}`;
      return [content,userIcon, displayAuthor, results];

    } catch (e) {
      // set state to returns error if an error occurred
      this.setState({
        error: "No document found for this ID!",
        loading: false
      });
      // log the error to console
      Log.error("SpfxDocumentDetail", e.message);
    }
  }
}