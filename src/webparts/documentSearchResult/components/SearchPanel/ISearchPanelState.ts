import { ISearchPanelValues } from "./ISearchPanelValues";
export interface ISearchPanelState {
    functionOptions: any[];
    locationOptions: any[];
    categoryOptions: any[];
    values?: ISearchPanelValues;
    error?: string;
}