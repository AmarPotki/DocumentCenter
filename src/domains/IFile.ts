/**
 * File item interface
 * Create File item to work with it internally
 */
export interface IFile {
  Id?: number;
  Title: string;
  Name: string;
  Size?: number;
  ServerRelativeUrl?: string;
}
