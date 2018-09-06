import { ITaxonomyValue } from "react-taxonomypicker";
import { Option } from "react-select";

export interface IDocumentDetailWebPartProps {
  searchResultPage: string;
  noDocumentMessage: string;
  showDescription: boolean;
  documentLibrary: string;
  shareBodyContent: string;
  // createdDateLabel:string;
  // issueDateLabel:string;
  contentManageProperty: string;
  managedProperties: string;
  managedPropertyForDescription: string;
  defaultValue?: ITaxonomyValue | ITaxonomyValue[] | Option | Option[] | string | string[] | number | number[] | boolean;
  defaultValueForDescription?: ITaxonomyValue | ITaxonomyValue[] | Option | Option[] | string | string[] | number | number[] | boolean;
}
