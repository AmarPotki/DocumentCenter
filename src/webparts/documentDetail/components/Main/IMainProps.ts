import { IDocumentService, ISearchService } from '../../../../services';
export interface IMainProps {
    documentPath: string;
    noDocumentsMessage: string;
    showDescription: boolean;
    documentService: IDocumentService;
    searchResultsPage: string;
    shareBodyContent: string;
    searchService: ISearchService;
    managedProperties: string;
    managedPropertyForDescription: string;
    webAbsoluteUrl: string;
    // createdDateLabel:string;
    // categoryLabel:string;
    // issueDateLabel:string;
}