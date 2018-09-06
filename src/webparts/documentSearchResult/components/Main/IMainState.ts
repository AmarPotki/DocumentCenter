import { ISearchOptions } from "../../../../domains";
import { ISearchPanelValues } from "../SearchPanel/ISearchPanelValues";
import { ISearchResultValues } from "../SearchResult/ISearchResultValues";

export interface IMainState {
    panelValues: ISearchPanelValues;
    resultValues: ISearchResultValues;
    isLoading: boolean;
    showPlaceholder: boolean;
    totalRows: number;
}