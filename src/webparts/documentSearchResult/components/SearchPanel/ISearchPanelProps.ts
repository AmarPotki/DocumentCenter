import { ITaxonomyHelper } from "../../../../common/ITaxonomyHelper";
import { IDocumentSearchResultWebPartProps } from "../../IDocumentSearchResultWebPartProps";
import { ISearchOptions } from "../../../../domains";
import { ISearchPanelValues } from "./ISearchPanelValues";
export interface ISearchPanelProps {
    taxonomyHelper: ITaxonomyHelper;
    webPartProperties: IDocumentSearchResultWebPartProps;
    onChange: (event: any) => void;
    onSearchInputKeyPress: (event: any) => void;
    onSearchButtonClicked: (event: any) => void;
    values: ISearchPanelValues;
}