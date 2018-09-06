import { IList, IPage } from "../domains";
import { SearchQueryBuilder, SearchResults } from "sp-pnp-js";
import { IRefinementValue } from "../domains/IRefinementValue";
import { ISearchSource } from "../domains/ISearchSource";

export interface ISearchService {
    /**
     * API to get the search results
     */
    GetSearchResult: (query: SearchQueryBuilder) => Promise<SearchResults>;
    GetAllManagedProperties(): Promise<Array<IRefinementValue>>;
    GetSearchSources(): Promise<Array<ISearchSource>>;
    GetNextPage(): Promise<SearchResults>;
    GetPrevPage(): Promise<SearchResults>;
}