import { IDocumentService, ISearchService } from '../../../../services';
export interface IAssociatedDocumentsProps {
    documentService: IDocumentService;
    primaryDocumentFolder: string;
    webUrl: string;
    shareBodyContent: string;
    managedPropertyForDescription: string;
    managedProperties: string;
    searchService: ISearchService;
    webAbsoluteUrl: string;
}