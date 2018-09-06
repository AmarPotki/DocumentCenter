import { IDocumentSearchResultWebPartProps } from "../../IDocumentSearchResultWebPartProps";
import { ISearchOptions } from "../../../../domains";
import { ISearchResultValues } from "./ISearchResultValues";
import { ISearchService, IDocumentService } from "../../../../services";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISearchResultProps {
    // context: WebPartContext;
    searchService: ISearchService;
    documentService: IDocumentService;
    webPartProperties: IDocumentSearchResultWebPartProps;
    values: ISearchResultValues;
    webAbsoluteUrl: string;
    onSortByChanged: (event: any) => void;
    loadMoreDocuments: () => void;
    clearSearch: () => void;
}