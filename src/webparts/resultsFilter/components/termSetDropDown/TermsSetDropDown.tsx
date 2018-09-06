import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IDropdownOption, Dropdown, DropdownMenuItemType } from 'office-ui-fabric-react/lib/';
import { ITermsSetDropDownState } from './ITermSetDropDownState';
import { ITermsSetDropDownProps } from './ITermSetDropDownProps';
import SPTermStorePickerService from '../../../../services/SPTermStorePickerService';
import ISPTermStorePickerService from '../../../../services/SPTermStorePickerService';
import { ITerm } from '../../../../services/ISPTermStorePickerService';

export default class TermsSetDropDown extends React.Component<ITermsSetDropDownProps, any> {
  private options: IDropdownOption[] = [];
  constructor(props: ITermsSetDropDownProps) {
    super(props);
    this.termStorService = this.props.termStorePickerService;
    this.state = {
      termSetName: `Filter By ${this.props.termSet.name}`,
      cleanTermSetId: this.termStorService._cleanGuid(this.props.termSet.id)
    };
    this.termStorService.getAllTerms(this.props.termSet.objectIdentity).then((terms: ITerm[]) => {
      if (terms !== null) {
        for (let term of terms) {
          this.options.push({ text: term.Name, key: term.Id });
        }
      }
    });
  }
  public render(): React.ReactElement<ITermsSetDropDownProps> {

    return (
        <Dropdown
          placeHolder={this.state.termSetName}
          label=""
          id="ddpType"
          ariaLabel=""
          onChanged={this._onChange}
          options={this.options}
        />
    );
  }
  private termStorService: ISPTermStorePickerService;
  private _onChange = (option: IDropdownOption): void => {
    let url: any = new URL(window.location.href);
    if (url.searchParams.has(this.props.termSet.name)) {
      url.searchParams.delete(this.props.termSet.name);
    }
    url.searchParams.set(this.props.termSet.name, option.text);
    if (history.pushState) {
      window.history.pushState({ path: url.href }, '', url.href);
    }
  }
}
