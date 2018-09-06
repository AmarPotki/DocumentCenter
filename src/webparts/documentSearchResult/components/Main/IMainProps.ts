import { IDocumentSearchResultWebPartProps } from "../../IDocumentSearchResultWebPartProps";
import { ISearchOptions } from "../../../../domains";
import { ISearchService, IDocumentService } from "../../../../services";
import { ITaxonomyHelper } from "../../../../common/ITaxonomyHelper";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMainProps {
    context: WebPartContext;
    searchService: ISearchService;
    documentService: IDocumentService;
    taxonomyHelper: ITaxonomyHelper;
    webPartProperties: IDocumentSearchResultWebPartProps;
    webAbsoluteUrl: string;
}