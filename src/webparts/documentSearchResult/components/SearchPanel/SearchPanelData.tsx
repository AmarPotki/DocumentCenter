import * as React from "react";
import { ISearchPanelProps } from "./ISearchPanelProps";
import { ISearchPanelState } from "./ISearchPanelState";
import { ITermGroup, ITermSet, ITerm, ITermStore } from "../../../../common/SPTaxonomyEntities";
import { SearchPanel } from "./SearchPanel";

export class SearchPanelData {
    constructor(private searchPanel: SearchPanel) {
    }
    public async GetFiltersOptions(): Promise<void> {

        let termStores: ITermStore[] = await this.searchPanel.props.taxonomyHelper.getTermStores();
        let termGroups: ITermGroup[] = await this.searchPanel.props.taxonomyHelper.getTermGroups(termStores[0].id);
        let termGroup: ITermGroup;
        let locationTermSet: ITermSet;
        let functionTermSet: ITermSet;
        let categoryTermSet: ITermSet;

        termGroups.forEach(tg => {
            if (tg.name === "BMI") {
                termGroup = tg;
                return;
            }
        });

        let termSets: ITermSet[] = await this.searchPanel.props.taxonomyHelper.getTermSets(termGroup);

        // this.searchPanel.state.categoryOptions.push(<option>{this.searchPanel.props.webPartProperties.categoryFilterLabel}</option>);
        // this.searchPanel.state.locationOptions.push(<option>{this.searchPanel.props.webPartProperties.locationFilterLabel}</option>);
        // this.searchPanel.state.functionOptions.push(<option>{this.searchPanel.props.webPartProperties.functionFilterLabel}</option>);

        for (let i: number = 0; i < termSets.length; i++) {
            let termSet: ITermSet = termSets[i];
            let termSetTitle: string = termSet.name;

            switch (termSetTitle) {
                case "Locations":
                    locationTermSet = termSet;
                    this.setFilterItems(termSet, this.searchPanel.state.locationOptions).then(() => {
                        this.searchPanel.setState({
                            error: ""
                        });
                    });
                    break;
                case "Functions":
                    functionTermSet = termSet;
                    this.setFilterItems(termSet, this.searchPanel.state.functionOptions).then(() => {
                        this.searchPanel.setState({
                            error: ""
                        });
                    });
                    break;
                case "Document Categories":
                    categoryTermSet = termSet;
                    this.setFilterItems(termSet, this.searchPanel.state.categoryOptions).then(() => {
                        this.searchPanel.setState({
                            error: ""
                        });
                    });
                    break;
            }
        }
    }

    public async setFilterItems(termSet: ITermSet, options: any[]): Promise<boolean> {
        let terms: ITerm[] = await this.searchPanel.props.taxonomyHelper.getTerms(termSet);
        for (let i: number = 0; i < terms.length; i++) {
            let term: ITerm = terms[i];
            options.push(<option className="opt-head">{term.labels[0].value}</option>);
            await this.getChildTerms(term, options);
        }
        return Promise.resolve(true);
    }

    public async getChildTerms(term: ITerm, options: any[]): Promise<boolean> {
        let childTerms: ITerm[] = await this.searchPanel.props.taxonomyHelper.getChildTerms(term);
        for (let i: number = 0; i < childTerms.length; i++) {
            let childTerm: ITerm = childTerms[i];
            options.push(<option className="opt-child">{childTerm.labels[0].value}</option>);
        }
        return Promise.resolve(true);
    }
}