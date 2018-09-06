import { ITaxonomyValue } from "react-taxonomypicker";
import { Option } from "react-select";

export interface IDocumentSearchResultWebPartProps {
  // categoryFilterLabel: string;
  // categoryLabel: string;
  // createdDateLabel: string;
  // defaultQueryText: string;
  documentDetailPage: string;
  // fileTypeFilterLabel: string;
  // functionFilterLabel: string;
  // issueDateLabel: string;
  listPath: string;
  // locationFilterLabel: string;
  // searchButtonText: string;
  // searchText: string;
  // shareBodyContent: string;
  // showCategoryFilter: boolean;
  showDescription: boolean;
  useDocumentDetailPage: boolean;
  // showFileTypeFilter: boolean;
  // showFunctionFilter: boolean;
  // showLocationFilter: boolean;
  showShareButton: boolean;
  enableSearchOrdering: boolean;
  // siteUrl: string;// not implemented yet , web want it to be a property for get lists by this url
  sortDocByLastModified: string;
  sortDocumentBy: string;
  managedProperties: string;
  managedPropertiesDefaultValue?: ITaxonomyValue | ITaxonomyValue[] | Option | Option[] | string | string[] | number | number[] | boolean;
  listsPathDefaultValue?: ITaxonomyValue | ITaxonomyValue[] | Option | Option[] | string | string[] | number | number[] | boolean;
}