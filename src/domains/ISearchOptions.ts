import { SortDirection } from "sp-pnp-js";
export interface ISearchOptions {
    queryText: string;
    sortProperty?: string;
    sortDirection?: SortDirection;
    rowLimit: number;
    refiners?: string;
    selectedProperties?: string[];
}