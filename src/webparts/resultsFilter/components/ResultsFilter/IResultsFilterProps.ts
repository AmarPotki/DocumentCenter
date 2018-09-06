import { ICheckedTermSets } from "../../../../controls/termSetPicker";
import ISPTermStorePickerService from '../../../../services/SPTermStorePickerService';
export interface IResultsFilterProps {
  terms: ICheckedTermSets;
  termStorePickerService: ISPTermStorePickerService;
  title: String;
}
