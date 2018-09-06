import { ICheckedTermSets } from './IPropertyFieldTermSetPicker';
import { ITermStore, IGroup, ITermSet, ITerm } from '../../services/ISPTermStorePickerService';
import { IPropertyFieldTermSetPickerPropsInternal } from './IPropertyFieldTermSetPicker';
import SPTermSetStorePickerService from '../../services/SPTermSetStorePickerService';

/**
 * PropertyFieldTermSetPickerHost properties interface
 */
export interface IPropertyFieldTermSetPickerHostProps extends IPropertyFieldTermSetPickerPropsInternal {

  onChange: (targetProperty?: string, newValue?: any) => void;
}

/**
 * PropertyFieldTermSetPickerHost state interface
 */
export interface IPropertyFieldTermSetPickerHostState {

  termStores?: ITermStore[];
  errorMessage?: string;
  openPanel?: boolean;
  loaded?: boolean;
  activeNodes?: ICheckedTermSets;
}

export interface ITermSetChanges {

  changedCallback: (termSet: ITermSet, checked: boolean) => void;
  activeNodes?: ICheckedTermSets;
}

export interface ITermGroupProps extends ITermSetChanges {

  group: IGroup;
  termstore: string;
  termsService: SPTermSetStorePickerService;
  multiSelection: boolean;
}

export interface ITermGroupState {

  expanded: boolean;
}

export interface ITermSetProps extends ITermSetChanges {

  termset: ITermSet;
  termstore: string;
  termsService: SPTermSetStorePickerService;
  autoExpand: () => void;
  multiSelection: boolean;
}

export interface ITermSetState {
  selected?: boolean;
  terms?: ITerm[];
  loaded?: boolean;
  expanded?: boolean;
}

export interface ITermProps extends ITermSetChanges {

  termset: string;
  term: ITerm;
  multiSelection: boolean;
}

export interface ITermState {

  selected?: boolean;
}
