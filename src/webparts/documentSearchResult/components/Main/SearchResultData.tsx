import * as React from "react";
import { IDocumentSearchResultWebPartProps } from "../../IDocumentSearchResultWebPartProps";
import { ISearchOptions } from "../../../../domains";
import { ISearchResponse } from "./ISearchResponse";
import { ISearchService } from "../../../../services";
import {
    SearchQuery,
    SearchQueryBuilder,
    SearchResults,
    Sort,
    SortDirection,
    QueryPropertyValueType
} from "sp-pnp-js";

export class SearchResultData {
    private documents: any[];
    constructor(private searchService: ISearchService, private props: IDocumentSearchResultWebPartProps) {
        this.LoadMoreDocuments = this.LoadMoreDocuments.bind(this);
    }

    public async GetSearchResults(options: ISearchOptions): Promise<SearchResults> {
        this.documents = [];
        let properties: string[];
        const appSearchSettings: SearchQuery = {
            TrimDuplicates: false,
            Querytext: "*",
            // tslint:disable-next-line:comment-format
            // SourceId: "8413cd39-2156-4e00-b54d-11efd9abdb89"
        };

        // let myproperties: Array<string> = ["Title", "Size", "Filename",
        //     "ContentType", "Created", "IssueDate", "DisplayAuthor",
        //     "FileExtension", "Path", "ServerRedirectedEmbedURL"];

        // let myproperties: Array<string> = [];
        // options.selectedProperties.forEach((key: string) => {
        //     myproperties.push(key);
        // });

        // tslint:disable-next-line:max-line-length
        this.props.managedProperties = "Title,Filename,FileExtension,Path,DisplayAuthor,Created,LastModifiedTime,ModifiedOWSDATE";

        // if (this.props.managedProperties.indexOf("HitHighlightedSummary") === -1) {
        //     this.props.managedProperties += ",HitHighlightedSummary";
        // }

        // if (this.props.managedProperties.indexOf("UniqueId") === -1) {
        //     this.props.managedProperties += ",UniqueId";
        // }

        // console.log("this.props.managedProperties: " + this.props.managedProperties);
        // if (options.selectedProperties.indexOf("UniqueId") === -1) {
        //     options.selectedProperties.push("UniqueId");
        // }

        let query: SearchQueryBuilder = SearchQueryBuilder
            .create(options.queryText, appSearchSettings)
            .properties({
                Name: "SourceName",
                Value: {
                    QueryPropertyValueTypeIndex: QueryPropertyValueType.StringType,
                    StrVal: this.props.listPath,
                }
            }, {
                Name: "SourceLevel",
                Value: {
                    QueryPropertyValueTypeIndex: QueryPropertyValueType.StringType,
                    StrVal: "SPSite"
                }
            })
            .rowLimit(options.rowLimit)
            .startRow(0)
            //     .selectProperties(...options.selectedProperties);
            // console.log("options.selectedProperties: " + options.selectedProperties);
            .selectProperties("Title", "Filename", "FileExtension", "Path", "DisplayAuthor",
            "UniqueId", "HitHighlightedSummary", "AuthorOWSUSER", "ContentType", "Created", "LastModifiedTime", "ModifiedOWSDATE");
        // .selectProperties("Title", "Size", "Filename",
        // "ContentType", "Created", "IssueDate", "DisplayAuthor",
        // "FileExtension", "Path", "ServerRedirectedEmbedURL", "UniqueId", "HitHighlightedSummary");

        if (options.refiners !== "" && options.refiners !== undefined) {
            query.refinementFilters(options.refiners);
        }
        if (options.sortProperty !== "" && options.sortProperty !== undefined) {
            let sort: Sort = { Property: options.sortProperty, Direction: options.sortDirection };
            query.sortList(sort);
        }
        this.searchService.GetSearchResult(query).then((value: SearchResults) => { console.log(value); }
        ).catch((err) => {
            console.log(err);
        });
        let searchResult: SearchResults = await this.searchService.GetSearchResult(query);
        return Promise.resolve(searchResult);
    }

    public async LoadMoreDocuments(): Promise<SearchResults> {
        let searchResult: SearchResults = await this.searchService.GetNextPage();
        return Promise.resolve(searchResult);
    }
}