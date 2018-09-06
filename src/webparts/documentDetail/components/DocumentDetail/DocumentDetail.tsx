import * as React from "react";
import styles from "../DocumentDetail.module.scss";
import { IDocumentDetailProps } from "./IDocumentDetailProps";
import { IDocumentDetailState } from "./IDocumentDetailState";
import { AssociatedDocuments } from "../AssociatedDocuments";
import { ICCDocument, ISearchOptions } from "../../../../domains";
import { Log } from "@microsoft/sp-core-library";
import { Dictionary, SearchQueryBuilder, SearchQuery, SearchResults } from "sp-pnp-js";
import { IUtilities, Utilities } from "../../../../common";
export class DocumentDetail extends React.Component<IDocumentDetailProps, IDocumentDetailState> {
  private utilities: IUtilities;
  /**
   * constructor
   * @param props: control properties
   */
  constructor(props: IDocumentDetailProps) {
    super(props);
    // bind default values to current state
    this.utilities = new Utilities();
    this.state = {
      isDocumentSet: false,
      loading: true,
      error: "",
      panelManagedProperties: [],
      document: {
        Id: 0, BMIDocumentCategory: "", ContentTypeName: "", Created: "", Name: "", ServerRelativeUrl: "",
        Title: "", Icon: "", UniqueId: "", Path: "", OnlinePath: "", DocumentSetFolder: "", IssueDate: "",
        fileType: '', displayAuthor: '', modified: '', userIcon: '',
        content: "", fields: new Dictionary<string>()
      }, results: null, webUrl: ''
    };
    this.getDocument(this.props.documentPath);
  }
  /**
   * Renders the control
   */
  public render(): JSX.Element {

    // replace URL of the property with document path and make a link for Share button
    let bodyContent: string = this.props.shareBodyContent;
    bodyContent = bodyContent.replace("##URL##", encodeURIComponent(this.state.document.Path));
    bodyContent = bodyContent.replace(/(\r\n|\n|\r)/gm, "%0D%0A");
    let mailTo: string = `mailto:?subject=${encodeURIComponent(this.state.document.Title)}&body=${bodyContent}`;

    return (
      <div>
        {
          this.state.error === "" &&
          <div>
            <div >
              <ul>

              </ul>
            </div>

            <div className={styles.hit}>
              <div className={styles.hitImage}>
                <img src={this.state.document.Icon} alt="icon" />
              </div>
              <div className={styles.hitContent}>
                <a className={styles.hitPrice}  ><img src={this.utilities.getIcon("Bookmark", false)} alt="Bookmark" /></a>
                <a className={styles.hitPrice} href={mailTo} ><img src={this.utilities.getIcon("Share", false)} alt="Share" /></a>
                <a className={styles.hitPrice}  ><img src={this.utilities.getIcon("Translate", false)} alt="Translate" /></a>
                <p className={styles.hitName}>
                  <a href={this.state.document.OnlinePath}>
                    {this.state.document.Title}
                  </a></p>
                <ul className={styles.documentMetaContainer}>
                  <li>
                    <strong>File Type: </strong> {this.state.document.fileType}</li>
                  <li>
                    <strong>Created: </strong> {this.state.document.Created}</li>
                  <li>
                    <strong>Modified: </strong> {this.state.document.modified}</li>
                </ul>

                <p className={styles.hitDescription}>
                  {this.utilities.stripHTMLFromText(this.state.document.content)}
                </p>
                <ul className={styles.documentMetaContainer}>
                  <li>
                    <img className={styles.user_icon}
                      src={this.state.document.userIcon} />
                    <strong className={styles.created_user}>Author: </strong> {this.state.document.displayAuthor}</li>
                  {(() => {
                    if (this.state.panelManagedProperties.length == 0) {
                      return;
                    }
                    let managedProperties: string[] = [];
                    managedProperties = this.props.managedProperties.split(",");
                    return (managedProperties.map((item: string) => {
                      {
                        if (this.state.results !== null) {
                          if (this.state.results.PrimarySearchResults[0][item] !== "" && item !== "") {
                            return (<li><strong>{item}: </strong>
                              {this.state.results.PrimarySearchResults[0][item]}   </li>);
                          }

                        }
                      }

                    }));
                  })()}
                </ul>
              </div>
            </div>
            {
              // render associated documents
              (this.state.isDocumentSet == true &&

                <AssociatedDocuments webUrl={this.state.webUrl} primaryDocumentFolder={this.state.document.DocumentSetFolder} {...this.props} />)
            }
          </div>
        }
        {
          // display error
          this.state.error !== "" &&
          <div>{this.state.error}</div>
        }
      </div>
    );
  }
  /**
   * Get Document information
   */
  private async getDocument(uniqueId: string): Promise<void> {
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
      if (results.PrimarySearchResults[0]["FileExtension"] !== null) {
        // call getDocumentByPath to get document information
        let doc: ICCDocument = await this.props.documentService.getDocumentByPath(documentPath, results.PrimarySearchResults[0]["SPWebUrl"]);
        doc.content = results.PrimarySearchResults[0][this.props.managedPropertyForDescription];
        doc.displayAuthor = results.PrimarySearchResults[0]["DisplayAuthor"];
        doc.fileType = results.PrimarySearchResults[0]["FileType"];
        doc.modified = results.PrimarySearchResults[0]["ModifiedOWSDATE"];
        doc.userIcon = `${this.props.webAbsoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${results.PrimarySearchResults[0]["AuthorOWSUSER"].split("|")[0].trim()}`;
        this.setState({
          loading: false,
          panelManagedProperties: managedProperties,
          document: doc,
          results: results,
          error: ""
        });
      } else {
        let webUrl = results.PrimarySearchResults[0]["SPWebUrl"];
        let doc = {
          Id: 0, BMIDocumentCategory: "", ContentTypeName: "", Created: "", Name: results.PrimarySearchResults[0]["Title"]
          , ServerRelativeUrl: "", fileType: '', displayAuthor: '', modified: '',
          Title: results.PrimarySearchResults[0]["Title"], Icon: this.utilities.getIcon("Folder", true),
          UniqueId: uniqueId, Path: "", OnlinePath: "", DocumentSetFolder: documentPath,
          IssueDate: "", userIcon: '',
          content: "", fields: new Dictionary<string>()
        };
        doc.content = results.PrimarySearchResults[0][this.props.managedPropertyForDescription];
        doc.displayAuthor = results.PrimarySearchResults[0]["DisplayAuthor"];
        doc.fileType = results.PrimarySearchResults[0]["FileType"];
        doc.modified = results.PrimarySearchResults[0]["ModifiedOWSDATE"];
        doc.userIcon = `${this.props.webAbsoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${results.PrimarySearchResults[0]["AuthorOWSUSER"].split("|")[0].trim()}`;
        doc.Created = doc.modified = results.PrimarySearchResults[0]["Created"];
        this.setState({
          isDocumentSet: true,
          loading: false,
          panelManagedProperties: managedProperties,
          document: doc,
          results: results,
          webUrl: webUrl,
        });
      }
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