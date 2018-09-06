import * as React from "react";
import styles from "../DocumentSearchResult.module.scss";
import { DisplayMode } from "@microsoft/sp-core-library";
import { ISearchResultProps } from "./ISearchResultProps";
import { ISearchResultState } from "./ISearchResultState";
import { IUtilities, Utilities } from "../../../../common";
import { SearchResults } from "sp-pnp-js";
import { unstable_renderSubtreeIntoContainer } from "react-dom";
const loadingIcon: any = require("../../../../assets/loading.gif");

export class SearchResult extends React.Component<ISearchResultProps, ISearchResultState> {
  private utilities: IUtilities;
  private associatedDocsArray: any[];
  private tenantUrl: string = `${window.location.protocol}//${window.location.host}`;
  constructor(props: ISearchResultProps) {
    super(props);
    this.utilities = new Utilities();
    this.associatedDocsArray = [];
    this.state = {
      error: "",
      loading: true
    };
  }

  public shareDocument(path: string, title: string): void {
    // let bodyContent: string = this.props.webPartProperties.shareBodyContent;
    // bodyContent = bodyContent.replace("##URL##", encodeURIComponent(path));
    // bodyContent = bodyContent.replace(/(\r\n|\n|\r)/gm, "%0D%0A");
    // let mailTo: string = `mailto:?subject=${encodeURIComponent(title)}&body=${bodyContent}`;
    // window.open(mailTo, "_self");
  }

  public documentDetail(path: string, uniqueId: string): void {
    let url: string = (this.props.webPartProperties.useDocumentDetailPage) ?
      // tslint:disable-next-line:max-line-length
      `${this.props.webAbsoluteUrl}/SitePages/${this.props.webPartProperties.documentDetailPage}?UniqueId=${uniqueId}&Source=${encodeURIComponent(window.location.toString())}`
      : path;

    window.open(url, "_self");

    // let url: string = (this.props.webPartProperties.documentDetailPage === undefined ||
    //   this.props.webPartProperties.documentDetailPage === null ||
    //   this.props.webPartProperties.documentDetailPage === "") ? path :
    // tslint:disable-next-line:max-line-length
    //   `${this.props.webAbsoluteUrl}/SitePages/${this.props.webPartProperties.documentDetailPage}?UniqueId=${uniqueId}&Source=${encodeURIComponent(window.location.toString())}`;
    // tslint:disable-next-line:max-line-length
    // // let url: string = `${this.props.webAbsoluteUrl}/SitePages/${this.props.webPartProperties.documentDetailPage}?DocumentPath=${path}&Source=${encodeURIComponent(window.location.toString())}`;
    // window.open(url, "_self");
  }

  public AddDocumentsToResult(results: SearchResults): any[] {
    let documents: any[] = [];

    results.PrimarySearchResults.forEach((result: any) => {
      // // tslint:disable-next-line:no-string-literal
      // let serverRelativeUrl: string = result["Path"].replace(this.tenantUrl, "");
      // // tslint:disable-next-line:no-string-literal
      // let isCCDocument: boolean = result["ContentType"].indexOf("CCDocument") !== -1;
      // // tslint:disable-next-line:no-string-literal
      // let isCCAssociatedDocument: boolean = result["ContentType"].indexOf("CCAssociatedDocument") !== -1;
      // let isPrimaryDocument: boolean = serverRelativeUrl.split("/").length > 3;
      // let folderServerRelatedUrl: string = serverRelativeUrl.substr(0, serverRelativeUrl.lastIndexOf("/"));

      // // tslint:disable-next-line:no-string-literal
      // if (result["Path"].indexOf("DispForm.aspx") === -1) {
      //   // tslint:disable-next-line:no-string-literal
      //   // console.log(result["Path"]);

      //   if ((isCCDocument && isPrimaryDocument) || isCCAssociatedDocument) {
      //     if (this.associatedDocsArray.indexOf(folderServerRelatedUrl) === -1 &&
      //       folderServerRelatedUrl.toLowerCase() !== "/sites/bmidocuments/shared documents") {
      //       this.associatedDocsArray.push(folderServerRelatedUrl);

      //       this.getPrimaryDocumentElement(result, folderServerRelatedUrl).then(p => {
      //         documents.push(p);
      //       });
      //     }
      //   } else {
      //     // tslint:disable-next-line:no-string-literal
      //     // console.log(result["Path"]);
      //     documents.push(this.getCCDocumentElement(result));
      //   }
      // }

      documents.push(this.getCCDocumentElement(result));
    });
    return documents;
  }

  public getCCDocumentElement(result: any): JSX.Element {
    {/* tslint:disable-next-line:no-string-literal */ }
    let serverRelativeUrl: string = result["Path"].replace(this.tenantUrl, "");
    return (
      <div className={styles.hit}>
        <div className={styles.hit_image}>
          <img onClick={() => this.documentDetail(serverRelativeUrl, result["UniqueId"])}
            src={this.utilities.getDocumentIcon(result["FileExtension"], true)}
            alt={result["Filename"]} title={result["Filename"]} />
        </div>
        <div className={styles.hit_content}>
          <a className={styles.hit_price}><img src={this.utilities.getIcon("Bookmark", false)} alt="Bookmark" /></a>
          <a className={styles.hit_price}><img src={this.utilities.getIcon("Share", false)} alt="Share" /></a>
          <a className={styles.hit_price}><img src={this.utilities.getIcon("Translate", false)} alt="Translate" /></a>
          <h2 className={styles.hit_name}>
            <a onClick={() => this.documentDetail(serverRelativeUrl, result["UniqueId"])}>
              {result["Title"]}
            </a></h2>
          <ul className={styles.documentMetaContainer}>
            {(() => {
              let managedProperties: string[] = [];
              managedProperties = this.props.webPartProperties.managedProperties.split(",");
              return (managedProperties.map((item: string) => {
                {
                  if ((result !== null && item !== "HitHighlightedSummary" && item !== "Path" && item !== "UniqueId")) {
                    return (<li><strong>{item}: </strong> {result[item]}   </li>);
                  }
                }
              }));
            })()}
          </ul>
          {/* tslint:disable-next-line:max-line-length */}
          <p className={styles.hit_description}>
            {this.props.webPartProperties.showDescription &&
              this.utilities.stripHTMLFromText(result["HitHighlightedSummary"])}
          </p>
          {/* tslint:disable-next-line:max-line-length */}
          <img className={styles.user_icon} src={`${this.props.webAbsoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${(result["AuthorOWSUSER"] === undefined || result["AuthorOWSUSER"] === null || result["AuthorOWSUSER"] === "") ? "" : result["AuthorOWSUSER"].split("|")[0].trim()}`} />
          {/* tslint:disable-next-line:max-line-length */}
          <span className={styles.created_user}>Created by: {(result["AuthorOWSUSER"] === undefined || result["AuthorOWSUSER"] === null || result["AuthorOWSUSER"] === "") ? "Unknown" : result["AuthorOWSUSER"].split("|")[1].trim()}</span>
        </div>
      </div>
    );
  }

  public async getPrimaryDocumentElement(result: any, folderUrl: string): Promise<JSX.Element> {
    let documentSetDocuments: any[] = await this.props.documentService.getDocumentSetDocuments(folderUrl);
    let primaryDocument: any;
    let associatedDocuments: any[] = [];
    documentSetDocuments.forEach(doc => {
      if (doc.ListItemAllFields.ContentType.Name === "CCDocument") {
        primaryDocument = doc;
      } else {
        let fileExtension: string = doc.Name.split(".").pop();
        associatedDocuments.push(
          <a href={doc.ServerRelativeUrl} target="_blank">
            <div className="associatedDocContainer">
              <img src={this.utilities.getDocumentIcon(fileExtension, false)} title={doc.Name} />
              <span> ${doc.Name}</span>
            </div>
          </a>);
      }
    });
    // let fileExtension: string = primaryDocument.Name.split(".").pop();
    // let path: string = this.tenantUrl + primaryDocument.ServerRelativeUrl;
    // let primaryDocumentName: string = primaryDocument.Name.substr(0, primaryDocument.Name.lastIndexOf("."));
    // let category: string = primaryDocument.ListItemAllFields.BMI_x0020_Document_x0020_Category != null ?
    //   primaryDocument.ListItemAllFields.BMI_x0020_Document_x0020_Category.Label : "Unknow";


    let document: JSX.Element = (
      <div></div>
      //   <div className="search-result-item">
      //     <div className="search-result-downloads">
      //       <div className="text-center">
      //         <img src={this.utilities.getDocumentIcon(fileExtension, true)} title={primaryDocument.Name} />
      //       </div>
      //       <div className="no-downloads"> <br />
      //         <div className="search-result-details">
      //           <div className="search-result-header">
      //             <div className="row">
      //               <div className="col-md-9 col-lg-9 col-sm-12 col-xs-12">
      //                 <h3><a href={path} title={primaryDocument.ListItemAllFields.Title}>{primaryDocumentName}</a></h3>
      //                 <ul className="search-result-meta">
      //                   <li>
      //                     <strong>{this.props.webPartProperties.categoryLabel}: </strong>
      //                     {category}
      //                   </li>
      //                   <li>
      //                     <strong>{this.props.webPartProperties.createdDateLabel}: </strong>
      //                     {this.utilities.getDateFromISOString(result.Created)}
      //                   </li>
      //                   {
      //                     primaryDocument.ListItemAllFields.IssueDate &&
      //                     <li>
      //                       <strong>{this.props.webPartProperties.issueDateLabel}: </strong>
      //                       {this.utilities.getDateFromISOString(primaryDocument.ListItemAllFields.IssueDate)}
      //                     </li>
      //                   }
      //                 </ul>
      //               </div>
      //               <div className="col-sm-12 col-xs-12 col-md-3 col-lg-3 search-result-button-container">
      //                 {/* tslint:disable-next-line:no-string-literal */}
      //                 <a href="#" onClick={() => this.documentDetail(primaryDocument.ServerRelativeUrl, result["UniqueID"])}
      //                   className="search-result-button">Show Details</a>
      //                 {
      //                   this.props.webPartProperties.showShareButton &&
      // tslint:disable-next-line:max-line-length
      //                   <a href="#" onClick={() => this.shareDocument(path, primaryDocumentName)} className="search-result-button">Share</a>
      //                 }
      //               </div>
      //             </div>
      //           </div>
      //           {
      //             this.props.webPartProperties.showDescription &&
      //             <div className="search-result-description">{primaryDocument.ListItemAllFields.KpiDescription}</div>
      //           }
      //           <div className="search-result-footer">
      //             <ul className="search-result-meta-footer">
      //               <li><strong>Associated Document(s):</strong></li>
      //               <li>
      //                 {associatedDocuments}
      //               </li>
      //             </ul>
      //           </div>
      //         </div>
      //       </div>
      //     </div>
      //   </div>
    );

    return Promise.resolve(document);
  }

  public render(): JSX.Element {
    return (
      <div>
        {
          this.props.values.searchResults !== undefined &&
          <div className={styles.documentSearchResult}>
            <div className={styles.hit}>
              {this.props.webPartProperties.enableSearchOrdering &&
                <select value={this.props.values.sortBy} id="sel-sortby"
                  onChange={this.props.onSortByChanged} className={styles.searchSortbyDropdown}>
                  <option value="Relevance">Sort by Relevance</option>
                  <option value="Created">Sort by Created date</option>
                  <option value="Modified">Sort by Modified Date</option>
                </select>}
            </div>
            <div>
              {
                this.AddDocumentsToResult(this.props.values.searchResults)
              }
            </div>
            <div className={styles.loadmore}>
              {
                this.props.values.searchResults.TotalRows > 0 &&
                <a onClick={this.props.loadMoreDocuments}>
                  Load more (found {this.props.values.searchResults.TotalRows}) <img src={this.utilities.getIcon("DropDown", false)} />
                </a>
              }
            </div>
          </div>
        }
      </div>
    );
  }
}