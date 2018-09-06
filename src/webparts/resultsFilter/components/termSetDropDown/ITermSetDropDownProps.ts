import { ICheckedTermSet } from "../../../../controls/termSetPicker";
import ISPTermStorePickerService from '../../../../services/SPTermStorePickerService';
export interface ITermsSetDropDownProps {
    termSet: ICheckedTermSet;
    termStorePickerService: ISPTermStorePickerService;

}