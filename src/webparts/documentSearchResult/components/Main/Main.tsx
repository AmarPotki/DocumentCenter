import * as React from "react";
import styles from "../DocumentSearchResult.module.scss";
import { FilterIdentities, SearchSchema } from "../../../../common";
import { IMainProps } from "./IMainProps";
import { IMainState } from "./IMainState";
import { ISearchOptions } from "../../../../domains";
import { ISearchPanelValues } from "../SearchPanel/ISearchPanelValues";
import { SearchPanel } from "../SearchPanel";
import { SearchResult } from "../SearchResult";
import { SearchResultData } from "./SearchResultData";
import { SortDirection, log } from "sp-pnp-js";
import { Spinner, SpinnerSize } from "office-ui-fabric-react";
import { Placeholder } from "@pnp/spfx-controls-react/lib/PlaceHolder";
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";

export class Main extends React.Component<IMainProps, IMainState> {
    private searchResultData: SearchResultData;
    private path: string;
    private tmpQueryText: string;
    private oldURL: string;
    private oldHash: string;

    constructor(props: IMainProps) {
        super(props);

        // tslint:disable-next-line:max-line-length
        // (path:"" *) OR (path:"" *)

        // let tmpListPaths: string[] = this.props.webPartProperties.listPath.split(",");

        // for (let listPath of tmpListPaths) {
        //     // this.path
        // tslint:disable-next-line:max-line-length
        //     this.tmpQuertyText = `(path:\"${this.props.webAbsoluteUrl}/${this.props.webPartProperties.listPath}\" ${this.props.webPartProperties.defaultQueryText}) OR `;
        // }

        // tslint:disable-next-line:max-line-length
        // this.tmpQuertyText = "(path:\" Documents\" *) OR (path:\"" *)";

        // console.log("this.tmpQuertyText: ");
        // console.log(this.tmpQuertyText);

        // let tmpProps: string[] = [];

        // tmpProps = this.props.webPartProperties.managedProperties.split(",");

        // if (this.props.webPartProperties.managedProperties === undefined) {
        //     this.props.webPartProperties.managedProperties = "";
        // }

        // if (this.props.webPartProperties.managedProperties instanceof String) {
        //     tmpProps.push(this.props.webPartProperties.managedProperties);
        // } else {
        //     this.props.webPartProperties.managedProperties.map((item) => {
        //         tmpProps.push(item);
        //     });
        // }

        // let initSearchOptions: ISearchOptions = {
        //     // queryText: `${this.path} ${this.props.webPartProperties.defaultQueryText}`,
        //     queryText: this.tmpQuertyText,
        //     rowLimit: 10,
        //     sortDirection: SortDirection.Descending,
        //     sortProperty: "Created",
        //     selectedProperties: tmpProps
        // };

        this.state = {
            panelValues: {
                categoryDropDown: "",
                functionDropDown: "",
                fileTypeDropDown: "",
                locationDropDown: "",
                searchInput: ""
            },
            resultValues: {
                sortBy: "Created",
                searchResults: undefined
            },
            isLoading: false,
            showPlaceholder: (this.props.webPartProperties.listPath === null ||
                this.props.webPartProperties.listPath === ""),
            totalRows: 0
        };

        this.onSearchPanelChanged = this.onSearchPanelChanged.bind(this);
        this.loadMoreDocuments = this.loadMoreDocuments.bind(this);
        this.searchByFilter = this.searchByFilter.bind(this);
        this.onSortByChanged = this.onSortByChanged.bind(this);
        this.clearSearch = this.clearSearch.bind(this);
        this.onSearchInputKeyPress = this.onSearchInputKeyPress.bind(this);
        // this.searchResultData = new SearchResultData(this.props.searchService, this.props.webPartProperties);
        // this.searchResultData.GetSearchResults(initSearchOptions).then(response => {
        //     this.setState({
        //         resultValues: {
        //             searchResults: response,
        //             sortBy: "Created"
        //         }
        //     });
        // });

        let location: Location = window.location;
        this.oldURL = location.href;
        this.oldHash = location.hash;
        // check the location hash on a 100ms interval
        setInterval(() => {
            let newURL: string = window.location.href,
                newHash: string = window.location.hash;
            // if the hash has changed and a handler has been bound...
            if (newURL !== this.oldURL) {// && typeof window.onhashchange === "function") {
                // execute the handler
                // window.onhashchange({
                //     type: "hashchange",
                //     oldURL: oldURL,
                //     newURL: newURL,
                //     bubbles: false,
                //     cancelable: false,
                //     cancelBubble: false,
                //     currentTarget: event
                // });

                this.locationHashChanged(newURL);
                // console.log("Loction Changed, New Url: " + newURL, " Old Url: " + this.oldURL);
                this.oldURL = newURL;
                this.oldHash = newHash;
            }
        }, 1000);
    }

    public locationHashChanged(newUrl: string): void {
        // let queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
        let queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(newUrl);

        if (queryParameters.getValue("k") !== undefined) {
            this.state.panelValues.searchInput =
                queryParameters.getValue("k") === undefined ? "" : queryParameters.getValue("k");

            this.setState({
                panelValues: this.state.panelValues
            });

            this.searchByFilter();
        }
    }

    public componentDidMount(): void {
        // this.context.window.onhashchange = this.locationHashChanged.bind(this);
        // window.onhashchange = this.locationHashChanged.bind(this);

        let queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);

        if (queryParameters.getValue("k") !== undefined) {
            this.state.panelValues.searchInput =
                queryParameters.getValue("k") === undefined ? "" : queryParameters.getValue("k");

            this.setState({
                panelValues: this.state.panelValues
            });

            this.searchByFilter();
        }
    }

    public onSearchPanelChanged(event: any): void {
        let controlId: any = event.target.id;
        let value: string = event.target.value;
        this.state.panelValues[controlId] = value;
        this.setState({
            panelValues: this.state.panelValues
        });
        if (event.target.type === "select-one") {
            this.searchByFilter();
        }
    }

    public onSortByChanged(event: any): void {
        this.state.resultValues.sortBy = event.target.value;
        this.setState({
            resultValues: this.state.resultValues
        });
        this.searchByFilter();
    }

    public clearSearch(): void {
        this.setState({
            resultValues: {
                searchResults: undefined,
                sortBy: "Created"
            },
            panelValues: {
                categoryDropDown: "",
                fileTypeDropDown: "",
                functionDropDown: "",
                locationDropDown: "",
                searchInput: ""
            }
        });
    }

    public onSearchInputKeyPress(event: any): void {
        if (event.charCode === 13) {
            event.preventDefault();
            event.stopPropagation();
            this.searchByFilter();
        }
    }

    public searchByFilter(): void {
        // if (this.state.showPlaceholder) {
        //     return;
        // }

        this.setState({
            isLoading: true
        });

        let searchFilters: Array<string> = [];
        let refiner: string = "";
        let sortByProperty: string = "";

        let tmpListPaths: string[] = this.props.webPartProperties.listPath.split(",");
        this.tmpQueryText = "";

        for (let listPath of tmpListPaths) {
            // tslint:disable-next-line:max-line-length
            // this.tmpQueryText += `(path:\"${this.props.webAbsoluteUrl}/${listPath}\" ${this.props.webPartProperties.defaultQueryText}) OR `;
            this.tmpQueryText += `(path:\"${this.props.webAbsoluteUrl}/${listPath}\") OR `;
        }

        this.tmpQueryText = this.tmpQueryText.substring(0, this.tmpQueryText.lastIndexOf("OR") - 1).trim();

        // searchFilters.push(this.path);
        // - searchFilters.push(this.tmpQueryText);
        // searchFilters.push(this.props.webPartProperties.defaultQueryText);
        if (this.state.panelValues[FilterIdentities.SearchInput] !== "") {
            searchFilters.push(this.state.panelValues[FilterIdentities.SearchInput]);
        }
        // if (this.state.panelValues[FilterIdentities.CategoryDropDown] !== "" &&
        //     this.state.panelValues[FilterIdentities.CategoryDropDown] !== this.props.webPartProperties.categoryFilterLabel) {
        //     searchFilters.push(`${SearchSchema.Category}:${this.state.panelValues[FilterIdentities.CategoryDropDown]}`);
        // }
        // if (this.state.panelValues[FilterIdentities.FunctionDropDown] !== "" &&
        //     this.state.panelValues[FilterIdentities.FunctionDropDown] !== this.props.webPartProperties.functionFilterLabel) {
        //     searchFilters.push(`${SearchSchema.Function}:${this.state.panelValues[FilterIdentities.FunctionDropDown]}`);
        // }
        // if (this.state.panelValues[FilterIdentities.LocationDropDown] !== "" &&
        //     this.state.panelValues[FilterIdentities.LocationDropDown] !== this.props.webPartProperties.locationFilterLabel) {
        //     searchFilters.push(`${SearchSchema.Location}:${this.state.panelValues[FilterIdentities.LocationDropDown]}`);
        // }
        // if (this.state.panelValues[FilterIdentities.FileTypeDropDown] !== "" &&
        //     this.state.panelValues[FilterIdentities.FileTypeDropDown] !== this.props.webPartProperties.fileTypeFilterLabel) {
        //     refiner = `fileExtension:equals("${this.state.panelValues[FilterIdentities.FileTypeDropDown]}")`;
        // }
        if (this.state.resultValues.sortBy !== "Relevance") {
            sortByProperty = this.state.resultValues.sortBy;
        }

        let tmpProps: string[] = [];

        tmpProps = this.props.webPartProperties.managedProperties.split(",");

        // if (this.props.webPartProperties.managedProperties === undefined) {
        //     this.props.webPartProperties.managedProperties = [""];
        // }

        // if (this.props.webPartProperties.managedProperties instanceof String) {
        //     tmpProps.push(this.props.webPartProperties.managedProperties);
        // } else {
        //     this.props.webPartProperties.managedProperties.map((item) => {
        //         tmpProps.push(item);
        //     });
        // }

        let searchOptions: ISearchOptions = {
            queryText: searchFilters.join(" "),
            rowLimit: 10,
            refiners: refiner,
            sortProperty: sortByProperty,
            sortDirection: SortDirection.Descending,
            selectedProperties: tmpProps
        };

        this.searchResultData = new SearchResultData(this.props.searchService, this.props.webPartProperties);
        this.searchResultData.GetSearchResults(searchOptions).then(response => {
            this.setState({
                resultValues: {
                    searchResults: response
                },
                totalRows: response.TotalRows,
                isLoading: false
            });
        });

        this.render();
    }

    public loadMoreDocuments(): void {
        this.searchResultData.LoadMoreDocuments().then(response => {
            this.setState({
                resultValues: {
                    searchResults: response
                }
            });
        });
    }

    private _configureWebPart(): void {
        this.props.context.propertyPane.open();
    }

    public render(): JSX.Element {
        if (this.state.showPlaceholder) {
            return (
                <Placeholder
                    iconName="Edit"
                    iconText="List view web part configuration"
                    description="Please configure the web part before you can show the list view."
                    buttonLabel="Configure"
                    onConfigure={this._configureWebPart.bind(this)} />
            );
        }

        return (
            <div className={styles.documentSearchResult}>
                {/* <SearchPanel {...this.props} onChange={this.onSearchPanelChanged.bind(this)}
                    values={this.state.panelValues} onSearchButtonClicked={this.searchByFilter.bind(this)}
                    onSearchInputKeyPress={this.onSearchInputKeyPress.bind(this)} /> */}
                {
                    this.state.isLoading ?
                        (
                            <Spinner size={SpinnerSize.large} label="Retrieving Search Results ..." />
                        ) : (
                            this.state.totalRows === 0 ?
                                (
                                    <Placeholder
                                        iconName="InfoSolid"
                                        iconText="No items found"
                                        description="The list or library you selected does not contain items." />
                                ) : (
                                    <div>
                                        <SearchResult {...this.props} values={this.state.resultValues}
                                            onSortByChanged={this.onSortByChanged.bind(this)}
                                            loadMoreDocuments={this.loadMoreDocuments.bind(this)}
                                            clearSearch={this.clearSearch.bind(this)} />
                                    </div>
                                )
                        )
                }
            </div>
        );
    }
}
