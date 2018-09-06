import { IList, IPage, IResponseFile, ISearchOptions } from "../domains";
import { IListService } from "./IListService";
import { SearchQuery, SearchResults, SearchQueryBuilder } from "sp-pnp-js";
import * as pnp from "sp-pnp-js";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { ISearchService } from "./ISearchService";
import { IRefinementValue } from "../domains/IRefinementValue";
import { ResultTable } from "sp-pnp-js/lib/sharepoint/search";
import { ISearchSource } from "../domains/ISearchSource";

export class SearchService implements ISearchService {
  private currentResults: SearchResults = null;
  private page = 0;

  constructor(private context: IWebPartContext) {
    /**
     * Setup pnp to use current context
     */
    pnp.setup({
      spfxContext: this.context
    });
  }
  /**
   * API to get the search results
   */
  public async GetSearchResult(query: SearchQueryBuilder): Promise<SearchResults> {
    // reset the position
    this.page = 1;
    this.currentResults = await pnp.sp.search(query);
    return Promise.resolve(this.currentResults);
  }

  /**
   * API to get All Search Schema Managed Properties
   */
  public async GetAllManagedProperties(): Promise<Array<IRefinementValue>> {
    return new Promise<Array<IRefinementValue>>(async (resolve, reject) => {
      let refinementValues: Array<IRefinementValue> = new Array();
      let searchQuery: SearchQuery = {
        QueryTemplate: "*",
        Refiners: "managedproperties(filter=600/0/*)"
      };

      this.currentResults = await pnp.sp.search(searchQuery);
      if (this.currentResults.RawSearchResults.PrimaryQueryResult) {
        let refinementResultsRows: ResultTable = this.currentResults.RawSearchResults.PrimaryQueryResult.RefinementResults;
        // tslint:disable-next-line:no-string-literal
        let refinementRows: any = refinementResultsRows ? refinementResultsRows["Refiners"] : [];

        refinementRows.map((refiner) => {
          refiner.Entries.map((item) => {
            refinementValues.push(<IRefinementValue>{
              RefinementCount: item.RefinementCount,
              RefinementName: item.RefinementName,
              RefinementToken: item.RefinementToken,
              RefinementValue: item.RefinementValue
            });
          });
        });

        resolve(refinementValues);
      } else {
        reject(new Error("No Item(s) Found"));
      }
    });
  }

  /**
   * API to get All Search Sources
   */
  public async GetSearchSources(): Promise<Array<ISearchSource>> {
    return new Promise<Array<ISearchSource>>(async (resolve, reject) => {
      let searchSources: Array<ISearchSource> = new Array();

      searchSources.push(<ISearchSource>{ Id: "0", Name: "Document" });
      searchSources.push(<ISearchSource>{ Id: "1", Name: "Everything" });

      if (searchSources.length > 0) {
        resolve(searchSources);
      } else {
        reject(new Error("No Item(s) Found"));
      }
    });
  }

  /**
   * API to get the next page of current search results
   */
  public async GetNextPage(): Promise<SearchResults> {
    this.currentResults = await this.currentResults.getPage(++this.page);
    return Promise.resolve(this.currentResults);
  }

  /**
   * API to get the previous page of current search results
   */
  public async GetPrevPage(): Promise<SearchResults> {
    this.currentResults = await this.currentResults.getPage(--this.page);
    return Promise.resolve(this.currentResults);
  }
}