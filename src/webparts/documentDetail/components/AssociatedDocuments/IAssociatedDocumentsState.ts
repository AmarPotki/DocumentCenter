import { SearchResults } from "sp-pnp-js";
import { ICCDocument } from "../../../../domains";

export interface IAssociatedDocumentsState {
    associatedDocuments: any[];
    loading: boolean;
    error: string;
    documentsCount: number;
    panelManagedProperties: string[];
    results: SearchResults;
    content: string;
    displayAuthor: string;
}