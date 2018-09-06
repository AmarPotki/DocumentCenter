import { IFile } from "./IFile";
import { Dictionary } from "sp-pnp-js/lib/collections/collections";
/**
 * CCDocument interface
 */
export interface ICCDocument extends IFile {
    BMIDocumentCategory: string;
    Created: string;
    DocumentSetFolder: string;
    ServerRelativeUrl: string;
    ContentTypeId?: string;
    ContentTypeName: string;
    IssueDate: string;
    // primaryDocument:boolean;
    Icon: any;
    UniqueId: string;
    Path: string;
    OnlinePath: string;
    fields: Dictionary<string>;
    content: string;
    fileType: string;
    modified: string;
    displayAuthor: string;
    userIcon: string;
}