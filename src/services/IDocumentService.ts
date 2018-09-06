import { ICCDocument, IAssociatedDocument, IDocumentItem } from "../domains";
import { IServiceBase } from "./IServiceBase";
export interface IDocumentService extends IServiceBase {

  /**
   * API to get document properties by document ID
   */
  getDocumentById: (documentId: number, documentLibrary: string) => Promise<ICCDocument>;

  /**
   * API to get document properties by document path
   */
  getDocumentByPath: (documentPath: string, webUrl: string) => Promise<ICCDocument>;

  /**
   * API to get associated documents from a folder
   */
  getAssociatedDocuments: (folderName: string, webUrl: string) => Promise<IAssociatedDocument[]>;

  /**
   * API to get Documents from a library
   */
  getDocumentsByListName: (library: string, top: number, properties: string,
    order: string, ascending: boolean) => Promise<IDocumentItem[]>;

  /**
   * API to get Documents of a document set folder
   */
  getDocumentSetDocuments: (folderName: string) => Promise<any[]>;

  /**
   * API to get related documents from a library
   */
  getRelatedDocumentsByPath: (library: string, documentPath: string, top: number,
    orderby: string, ascending: boolean) => Promise<IDocumentItem[]>;
}