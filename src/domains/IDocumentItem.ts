import { IFile } from "./IFile";
/**
 * Document Item interface
 */
export interface IDocumentItem {
    DocIcon?: string;
    Title?: string;
    Icon?: any;
    Id: number;
    File: IFile;
    Created: string;
}