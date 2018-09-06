import { IDocumentService, ISearchService } from "../../../../services";
export interface IDocumentDetailProps {
  documentPath: string;
  showDescription: boolean;
  documentService: IDocumentService;
  searchService: ISearchService;
  shareBodyContent: string;
  managedProperties: string;
  managedPropertyForDescription: string;
  webAbsoluteUrl: string;
}
