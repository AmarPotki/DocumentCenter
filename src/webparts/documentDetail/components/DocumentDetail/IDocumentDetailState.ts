import { ICCDocument } from "../../../../domains";
import { SearchResults } from "sp-pnp-js";
export interface IDocumentDetailState {
    isDocumentSet: boolean;
    document: ICCDocument;
    loading: boolean;
    error: string;
    results: SearchResults;
    webUrl: string;
    panelManagedProperties: string[];
}