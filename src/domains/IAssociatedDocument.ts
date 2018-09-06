import { IFile } from "./IFile";
/**
 * AssociatedDocument interface
 */
export interface IAssociatedDocument extends IFile {
    Created: string;
    ServerRelativeUrl: string;
    ContentTypeId: string;
    Icon: any;
    Path: string;
    OnlinePath: string;
    UniqueId: string;
    TimeLastModified: string;
    FileType: string;
}