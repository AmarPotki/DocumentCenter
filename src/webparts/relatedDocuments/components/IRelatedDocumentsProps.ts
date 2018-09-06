import {IDocumentService} from '../../../services';
import {IRelatedDocumentsWebPartProps} from '../IRelatedDocumentsWebPartProps';
export interface IRelatedDocumentsProps {
  documentService:IDocumentService;
  webPartProps:IRelatedDocumentsWebPartProps;
  documentPath:string;
}
