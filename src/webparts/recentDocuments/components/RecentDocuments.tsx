import * as React from "react";
import styles from "./RecentDocuments.module.scss";
import { IRecentDocumentsProps } from "./IRecentDocumentsProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { Web, SearchQuery, SearchQueryBuilder, Sort, SortDirection, SearchResults, SearchResult, QueryPropertyValueType } from "sp-pnp-js";
import { ISearchService, SearchService, IGraphApiService, GraphApiService } from "../../../services";
import { IUtilities, Utilities } from "../../../common";
export interface IRecentDocumentsState {
  results: SearchResults;
  containerWidth: number;
}
export default class RecentDocuments extends React.Component<IRecentDocumentsProps, IRecentDocumentsState> {
  private utilities: IUtilities;
  private documents: any[];
  private searchService: ISearchService;
  private graphService: IGraphApiService;
  private timerID: any;
  constructor(props: IRecentDocumentsProps) {
    super(props);
    this.graphService = this.props.graphService;
    // this.graphService.GetRecentlyViewed();
    this.utilities = new Utilities();
    this.state = { results: null, containerWidth: 0 };
    this.documents = [];
    this.searchService = props.searchService;

    const appSearchSettings: SearchQuery = {
      TrimDuplicates: false,
      Querytext: "*"
    };

    let query: SearchQueryBuilder = SearchQueryBuilder
      .create("*", appSearchSettings).properties({
        Name: "SourceName",
        Value: {
          QueryPropertyValueTypeIndex: QueryPropertyValueType.StringType,
          StrVal: this.props.webPartProperties.sourceName,
        }
      }, {
        Name: "SourceLevel",
        Value: {
          QueryPropertyValueTypeIndex: QueryPropertyValueType.StringType,
          StrVal: "SPSite"
        }
      })
      .rowLimit(5)
      .startRow(0)
      .selectProperties("Title", "FileExtension", "Filename", "Created", "Path", "UniqueId");
    let sort: Sort = { Property: "Created", Direction: SortDirection.Descending };
    query.sortList(sort);
    this.searchService.GetSearchResult(query).then((value: SearchResults) => {
      this.setState({ results: value });
    }
    ).catch((err) => {
      console.log(err);
    });
  }
  private getOnlinePath(uniqueId: string, fileName: string, path: string, extension: string): string {
    let tenantUrl: string = window.location.protocol + "//" + window.location.host;
    return extension !== "pdf" ?
      `${this.props.webAbsoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${uniqueId}&file=${fileName}&action=default` :
      path;
  }

  public documentDetail(path: string, uniqueId: string, fileName: string, extension: string): void {
    let url: string = (this.props.webPartProperties.documentDetailPage === undefined ||
      this.props.webPartProperties.documentDetailPage === null ||
      this.props.webPartProperties.documentDetailPage === "") ? this.getOnlinePath(uniqueId, fileName, path, extension) :
      // tslint:disable-next-line:max-line-length
      `${this.props.webAbsoluteUrl}/SitePages/${this.props.webPartProperties.documentDetailPage}?UniqueId=${uniqueId}&Source=${encodeURIComponent(window.location.toString())}`;
    window.open(url, "_self");
  }
  public componentDidMount(): void {
    this.timerID = setInterval(
      () => this.watchSize(),
      1000
    );
  }
  protected watchSize(): void {
    const width: number = document.getElementById(styles.container).clientWidth;
    this.setState({ containerWidth: width });
  }
  public componentWillUnmount(): void {
    clearInterval(this.timerID);
  }

  public render(): React.ReactElement<IRecentDocumentsProps> {
    return (
      <div className={styles.recentDocuments} >
        <div className={`ms-Grid ${styles.container}`} id={styles.container}>
          <div className={`ms-Grid-row ${styles.row}`}>
            <div className={`ms-Grid-col ${styles.column}`}>
              <div className={styles.title}>{this.props.webPartProperties.title}</div>
            </div>
          </div>
          <ul className={`ms-Grid-row ${styles.row}`}>
            {(() => {
              if (this.state.results !== null) {
                return (this.state.results.PrimarySearchResults.map((item: SearchResult) => {
                  {
                    if (this.state.containerWidth > 420 && this.state.containerWidth < 800) {
                      // tslint:disable-next-line:max-line-length
                      return (<li className={`ms-Grid-col ms-lg6 ms-md6 ms-sm12 ${styles.column}`}>
                        <a onClick={() => this.documentDetail(item["UniqueId"],
                          item["Filename"], item["Path"], item["FileExtension"])}>
                          {/* tslint:disable-next-line:no-string-literal */}
                          <img src={this.utilities.getDocumentIcon(item["FileExtension"], false)} /></a>
                        {/* tslint:disable-next-line:max-line-length */}
                        <p><a onClick={() => this.documentDetail(item["UniqueId"],
                          item["Filename"], item["Path"], item["FileExtension"])} >{item["Title"]}</a>
                        </p><hr /></li>);
                    } else  if (this.state.containerWidth > 800) {
                      // tslint:disable-next-line:max-line-length
                      return (<li className={`ms-Grid-col ms-lg4 ms-md6 ms-sm12 ${styles.column}`}>
                        <a onClick={() => this.documentDetail(item["UniqueId"],
                          item["Filename"], item["Path"], item["FileExtension"])}>
                          {/* tslint:disable-next-line:no-string-literal */}
                          <img src={this.utilities.getDocumentIcon(item["FileExtension"], false)} /></a>
                        {/* tslint:disable-next-line:max-line-length */}
                        <p><a onClick={() => this.documentDetail(item["UniqueId"],
                          item["Filename"], item["Path"], item["FileExtension"])} >{item["Title"]}</a>
                        </p><hr /></li>);
                    }
                    else {
                      // tslint:disable-next-line:max-line-length
                      return (<li className={`ms-Grid-col ms-lg12 ms-md12 ms-sm12 ${styles.column}`}>
                        <a onClick={() => this.documentDetail(item["UniqueId"],
                          item["Filename"], item["Path"], item["FileExtension"])}>
                          {/* tslint:disable-next-line:no-string-literal */}
                          <img src={this.utilities.getDocumentIcon(item["FileExtension"], false)} /></a>
                        {/* tslint:disable-next-line:max-line-length */}
                        <p><a onClick={() => this.documentDetail(item["UniqueId"],
                          item["Filename"], item["Path"], item["FileExtension"])} >{item["Title"]}</a>
                        </p><hr /></li>);
                    }
                  }
                }
                ));
              }
            })()}
          </ul>
        </div>
      </div>
    );
  }
}
