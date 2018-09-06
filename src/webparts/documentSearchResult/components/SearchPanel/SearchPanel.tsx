import * as React from "react";
import { DisplayMode } from "@microsoft/sp-core-library";
import { ISearchPanelProps } from "./ISearchPanelProps";
import { ISearchPanelState } from "./ISearchPanelState";
import { ITermGroup, ITermSet, ITerm } from "../../../../common/SPTaxonomyEntities";
import { SearchPanelData } from "./SearchPanelData";
import { Log } from "@microsoft/sp-core-library";
import { SearchResult } from "../SearchResult/SearchResult";
import { ISearchOptions } from "../../../../domains";
import { FilterIdentities } from "../../../../common";


export class SearchPanel extends React.Component<ISearchPanelProps, ISearchPanelState> {
  constructor(props: ISearchPanelProps) {
    super(props);
    this.state = {
      functionOptions: [],
      locationOptions: [],
      categoryOptions: [],
      values: this.props.values
    };

    let taxonomyModel: SearchPanelData = new SearchPanelData(this);
    taxonomyModel.GetFiltersOptions();
  }

  public render(): JSX.Element {
    return (
      <div className="search-container">
        {/* <div className="search-filter-block">
          <h2>What Are You Looking For?</h2>
          <div className="search-input">
            <input type="text" id={FilterIdentities.SearchInput} onChange={this.props.onChange}
              value={this.props.values.searchInput} defaultValue={this.props.values.searchInput}
              onKeyPress={this.props.onSearchInputKeyPress} placeholder={this.props.webPartProperties.searchText} />
            <a href="#" onClick={this.props.onSearchButtonClicked}
              className="btn" id="search-button">{this.props.webPartProperties.searchButtonText}</a>
          </div>
          <label>And/Or</label>
          <div className="search-selects">
            <div className="row">
              <div className="${filterClassName} col-xs-6 ${!this.properties.showCategoryFilter ? 'hidden' : ''}">
                <select id={FilterIdentities.CategoryDropDown} onChange={this.props.onChange} value={this.props.values.categoryDropDown}>
                  {
                    this.state.categoryOptions.length < 1 &&
                    <option>Loading...</option>
                  }
                  {
                    this.state.categoryOptions
                  }
                </select>
              </div>
              <div className="${filterClassName} col-xs-6 ${!this.properties.showFunctionFilter ? 'hidden' : ''}">
                <select id={FilterIdentities.FunctionDropDown} onChange={this.props.onChange} value={this.props.values.functionDropDown}>
                  {
                    this.state.functionOptions.length < 1 &&
                    <option>Loading...</option>
                  }
                  {
                    this.state.functionOptions
                  }
                </select>
              </div>
              <div className="${filterClassName} col-xs-6 ${!this.properties.showLocationFilter ? 'hidden' : ''}">
                <select id={FilterIdentities.LocationDropDown} onChange={this.props.onChange} value={this.props.values.locationDropDown}>
                  {
                    this.state.locationOptions.length < 1 &&
                    <option>Loading...</option>
                  }
                  {
                    this.state.locationOptions
                  }
                </select>
              </div>
              <div className="${filterClassName} col-xs-6 ${!this.properties.showFileTypeFilter ? 'hidden' : ''}">
                <select id={FilterIdentities.FileTypeDropDown} onChange={this.props.onChange} value={this.props.values.fileTypeDropDown}>
                  <option value="Filter By File Type">Filter By File Type</option>
                  <option value="pdf">PDF</option>
                  <option value="xls">Excel</option>
                  <option value="doc">Word</option>
                  <option value="ppt">Powerpoint</option>
                </select>
              </div>
            </div>
          </div>
        </div> */}
      </div>
    );
  }
}