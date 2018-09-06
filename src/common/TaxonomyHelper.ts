/// <reference path="../common/SP.d.ts" />
import {
  IWebPartContext
} from "@microsoft/sp-webpart-base";
import {
  ITermStore,
  ITermSet,
  ITermGroup,
  ITerm
} from "./SPTaxonomyEntities";
import {
  ITaxonomyHelper
} from "./ITaxonomyHelper";
import { Label } from "office-ui-fabric-react";


/**
 * Interface for terms with path property and nested terns
 */
interface ITermWithTerms extends ITerm {
  path: string[];
  fullPath: string;
  terms?: ITerms;
}

/**
 * Interface that represents a map wit key term ID and value ITermWithTerms object
 */
interface ITerms {
  [name: string]: ITermWithTerms;
}

/**
 * Interface that represents a map with key term set ID and value ITerms object
 */
interface ITermSetTerms {
  [name: string]: ITerms;
}

/**
 * SharePoint Data Helper class.
 * Gets information from current web
 */
export class TaxonomyHelper implements ITaxonomyHelper {
  /**
   * cached term stores. This property can be changed to static to be able to use the same cache in different web parts
   */
  private _loadedTermStores: SPC.Taxonomy.ITermStoreCollection;
  /**
   * cached terms' hierarchy. This property can be changed to static to be able to use the same cache in different web parts
   */
  private _loadedTermsHierarchy: ITermSetTerms = {};
  /**
   * cached terms' flat list. This property can be changed to static to be able to use the same cache in different web parts
   */
  private _loadedTermsFlat: ITerms[] = [];


  /**
   * Web part context
   */
  public context: IWebPartContext;

  /**
   * ctor
   * @param context: web part context
   */
  public constructor(_context: IWebPartContext) {
    this.context = _context;
  }

  /**
   * API to get Term Stores
   */
  public getTermStores(): Promise<ITermStore[]> {
    return new Promise<ITermStore[]>((resolve) => {
      // term stores have been already loaded
      if (this._loadedTermStores) {
        // converting SPC.Taxonomy.ITermStore object to ITermStore objects
        const termStoreEntities: ITermStore[] = this.getTermStoreEntities(this._loadedTermStores);
        resolve(termStoreEntities);
        return;
      }

      //
      // need to load term stores
      //

      this.loadScripts().then(() => { // loading scripts first
        const taxonomySession: SPC.Taxonomy.TaxonomySession = this.taxonomySession;
        let termStores: SPC.Taxonomy.ITermStoreCollection = taxonomySession.get_termStores();
        this.clientContext.load(termStores);
        this.clientContext.executeQueryAsync(() => {
          // converting SPC.Taxonomy.ITermStore object to ITermStore objects
          const termStoreEntities: ITermStore[] = this.getTermStoreEntities(termStores);
          // caching loaded term stores
          this._loadedTermStores = termStores;
          resolve(termStoreEntities);
        }, () => {
          resolve([]);
        });

      });
    });
  }

  /**
   * API to get Term Groups by Term Store
   */
  public getTermGroups(termStoreId: string): Promise<ITermGroup[]> {

    return new Promise<ITermGroup[]>((resolve) => {
      this.getTermStoreById(termStoreId).then((termStore) => { // getting the term store
        if (!termStore) {
          resolve([]);
          return;
        }

        let groups: SPC.Taxonomy.ITermGroupCollection = termStore.get_groups();
        //
        // if Groups property is not loaded get_count will throw an error that will be handled to retrieve groups
        //
        try {

          if (!groups.get_count()) { // this will throw error if groups were not loaded
            resolve([]);
            return;
          }

          // converting SPC.Taxonomy.ITermGroup object to ITermGroup objects
          resolve(this.getTermGroupEntities(groups, termStore.get_id().toString()));
        } catch (ex) { // retrieving groups
          this.clientContext.load(groups);
          this.clientContext.executeQueryAsync(() => {
            // converting SPC.Taxonomy.ITermGroup object to ITermGroup objects
            resolve(this.getTermGroupEntities(groups, termStore.get_id().toString()));
          }, () => {
            resolve([]);
          });
        }
        // tslint:disable-next-line:no-empty
        finally { }
      });
    });
  }
  /**
   * API to get Term Sets by Term Group
   */
  public getTermSets(termGroup: ITermGroup): Promise<ITermSet[]> {
    return new Promise<ITermSet[]>((resolve) => {
      this.getTermStoreById(termGroup.termStoreId).then((termStore) => { // getting term store by id
        if (!termStore) {
          resolve([]);
          return;
        }

        this.getTermGroupById(termStore, termGroup.id).then((group) => { // getting term group by id
          if (!group) {
            resolve([]);
            return;
          }
          let termSets: SPC.Taxonomy.ITermSetCollection = group.get_termSets();
          //
          // if termSets property is not loaded get_count will throw an error that will be handled to retrieve term sets
          //
          try {
            if (!termSets.get_count()) { // this will throw error if term sets were not loaded
              resolve([]);
              return;
            }

            // converting SPC.Taxonomy.ITermSet object to ITermSet object
            resolve(this.getTermSetEntities(termSets, termGroup.id, termGroup.termStoreId));
          } catch (ex) { // retrieving term sets
            this.clientContext.load(termSets);
            this.clientContext.executeQueryAsync(() => {
              // converting SPC.Taxonomy.ITermSet object to ITermSet object
              resolve(this.getTermSetEntities(termSets, termGroup.id, termGroup.termStoreId));
            }, () => {
              resolve([]);
            });
          }
          // tslint:disable-next-line:no-empty
          finally { }
        });
      });
    });
  }
  /**
   * API to get Terms by Term Set
   */
  public getTerms(termSet: ITermSet): Promise<ITerm[]> {
    return new Promise<ITerm[]>((resolve) => {
      // checking if terms were previously loaded
      if (this._loadedTermsHierarchy[termSet.id]) {
        const termSetTerms: ITerms = this._loadedTermsHierarchy[termSet.id];
        // converting ITerms object to collection of ITerm objects
        resolve(this.getTermEntities(termSetTerms));
        return;
      }

      //
      // need to load terms
      //
      this.getTermStoreById(termSet.termStoreId).then((termStore) => { // getting store by id
        if (!termStore) {
          resolve([]);
          return;
        }

        this.getTermGroupById(termStore, termSet.termGroupId).then((group) => { // getting group by id
          if (!group) {
            resolve([]);
            return;
          }

          this.getTermSetById(termStore, group, termSet.id).then((set) => { // getting term set by id
            if (!set) {
              resolve([]);
              return;
            }

            let allTerms: SPC.Taxonomy.ITermCollection;
            //
            // if terms property is not loaded get_count will throw an error that will be handled to retrieve terms
            //
            try {
              allTerms = set.getAllTerms();

              if (!allTerms.get_count()) { // this will throw error if terms were not loaded
                resolve([]);
                return;
              }

              // caching terms
              this._loadedTermsHierarchy[termSet.id] = this.buildTermsHierarchy(allTerms, termSet.id);
              // converting ITerms object to collection of ITerm objects
              resolve(this.getTermEntities(this._loadedTermsHierarchy[termSet.id]));
            } catch (ex) { // retrieving terms
              this.clientContext.load(allTerms, "Include(Id, Name, Description, IsRoot, TermsCount, PathOfTerm, Labels)");
              this.clientContext.executeQueryAsync(() => {
                // caching terms
                this._loadedTermsHierarchy[termSet.id] = this.buildTermsHierarchy(allTerms, termSet.id);
                // converting ITerms object to collection of ITerm objects
                resolve(this.getTermEntities(this._loadedTermsHierarchy[termSet.id]));
              }, () => {
                resolve([]);
              });
            }
            // tslint:disable-next-line:no-empty
            finally { }

          });
        });
      });
    });
  }
  /**
   * API to get Terms by Term
   */
  public getChildTerms(term: ITerm): Promise<ITerm[]> {
    return new Promise<ITerm[]>((resolve) => {
      if (!this._loadedTermsFlat.length) {
        //
        // we can add logic to retrieve term from term Store
        // but I'll skip it for this example
        //
        resolve([]);
        return;
      }

      let terms: ITerms;
      // iterating through flat list of terms to find needed one
      for (let i: number = 0, len: number = this._loadedTermsFlat.length; i < len; i++) {
        const currTerm: ITermWithTerms = this._loadedTermsFlat[i][term.id];
        if (currTerm) {
          terms = currTerm.terms;
          break;
        }
      }

      if (!terms) {
        resolve([]);
        return;
      }

      // converting ITerms object to collection of ITerm objects
      resolve(this.getTermEntities(terms));

    });
  }

  /**
   * Loads scripts that are needed to work with taxonomy
   */
  private loadScripts(): Promise<void> {
    return new Promise<void>((resolve) => {
      //
      // constructing path to Layouts folder
      //
      let layoutsUrl: string = this.context.pageContext.site.absoluteUrl;
      if (layoutsUrl.lastIndexOf("/") === layoutsUrl.length - 1) {
        layoutsUrl = layoutsUrl.slice(0, -1);
      }

      layoutsUrl += "/_layouts/15/";

      this.loadScript(layoutsUrl + "init.js", "Sod").then(() => { // loading init.js
        resolve();
        // return this.loadScript(layoutsUrl + "SPC.runtime.js", "sp_runtime_initialize"); // loading SPC.runtime.js
      }).then(() => {
        resolve();
        // return this.loadScript(layoutsUrl + "SPC.js", "sp_initialize"); // loading SPC.js
      }).then(() => {
        resolve();
        // return this.loadScript(layoutsUrl + "SPC.taxonomy.js", "SPC.Taxonomy"); // loading SPC.taxonomy.js
      }).then(() => {
        resolve();
      });
    });
  }

  /**
   * Loads script
   * @param url: script src
   * @param globalObjectName: name of global object to check if it is loaded to the page
   */
  private loadScript(url: string, globalObjectName: string): Promise<void> {
    return new Promise<void>((resolve) => {
      let isLoaded: boolean = true;
      if (globalObjectName.indexOf(".") !== -1) {
        const props: string[] = globalObjectName.split(".");
        let currObj: any = window;

        for (let i: number = 0, len: number = props.length; i < len; i++) {
          if (!currObj[props[i]]) {
            isLoaded = false;
            break;
          }

          currObj = currObj[props[i]];
        }
      } else {
        isLoaded = !!window[globalObjectName];
      }
      // checking if the script was previously added to the page
      if (isLoaded || document.head.querySelector("script[src=\"" + url + "\"]")) {
        resolve();
        return;
      }

      // loading the script
      const script: HTMLScriptElement = document.createElement("script");
      script.type = "text/javascript";
      script.src = url;
      script.onload = () => {
        resolve();
      };
      document.head.appendChild(script);
    });
  }

  /**
   * Taxonomy session getter
   */
  private get taxonomySession(): SPC.Taxonomy.TaxonomySession {
    return SPC.Taxonomy.TaxonomySession.getTaxonomySession(this.clientContext);
  }

  /**
   * Client Context getter
   */
  private get clientContext(): SPC.ClientContext {
    return SPC.ClientContext.get_current();
  }

  /**
   * Converts SPC.Taxonomy.ITermStore objects to ITermStore objects
   * @param termStores: SPC.Taxonomy.ITermStoreCollection object
   */
  private getTermStoreEntities(termStores: SPC.Taxonomy.ITermStoreCollection): ITermStore[] {
    if (!termStores) {
      return [];
    }

    const termStoreEntities: ITermStore[] = [];
    for (let i: number = 0, len: number = termStores.get_count(); i < len; i++) {
      const termStore: SPC.Taxonomy.ITermStore = termStores.get_item(i);
      termStoreEntities.push({
        id: termStore.get_id().toString(),
        name: termStore.get_name()
      });
    }

    return termStoreEntities;
  }

  /**
   * Converts SPC.Taxonomy.ITermGroup objects to ITermGroup objects
   * @param termGroups: SPC.Taxonomy.ITermGroupCollection object
   * @param termStoreId: the identifier of parent term store
   */
  private getTermGroupEntities(termGroups: SPC.Taxonomy.ITermGroupCollection, termStoreId: string): ITermGroup[] {
    if (!termGroups) {
      return [];
    }
    const termGroupEntities: ITermGroup[] = [];
    for (let i: number = 0, len: number = termGroups.get_count(); i < len; i++) {
      const termGroup: SPC.Taxonomy.ITermGroup = termGroups.get_item(i);
      termGroupEntities.push({
        id: termGroup.get_id().toString(),
        termStoreId: termStoreId,
        name: termGroup.get_name(),
        description: termGroup.get_description()
      });
    }

    return termGroupEntities;
  }

  /**
   * Converts SPC.Taxonomy.ITermSet objects to ITermSet objects
   * @param termSets: SPC.Taxonomy.ITermSetCollection object
   * @param termGroupId: the identifier of parent term group
   * @param termStoreId: the identifier of parent term store
   */
  private getTermSetEntities(termSets: SPC.Taxonomy.ITermSetCollection, termGroupId: string, termStoreId: string): ITermSet[] {
    if (!termSets) {
      return [];
    }

    const termSetEntities: ITermSet[] = [];

    for (let i: number = 0, len: number = termSets.get_count(); i < len; i++) {
      const termSet: SPC.Taxonomy.ITermSet = termSets.get_item(i);
      termSetEntities.push({
        id: termSet.get_id().toString(),
        name: termSet.get_name(),
        description: termSet.get_description(),
        termGroupId: termGroupId,
        termStoreId: termStoreId
      });
    }

    return termSetEntities;
  }

  /**
   * Retrieves term store by its id
   * @param termStoreId: the identifier of the store to retrieve
   */
  private getTermStoreById(termStoreId: string): Promise<SPC.Taxonomy.ITermStore> {
    return new Promise<SPC.Taxonomy.ITermStore>((resolve) => {
      if (!this._loadedTermStores) { // term stores were not loaded, need to load them
        this.getTermStores().then(() => {
          return this.getTermStoreById(termStoreId);
        });
      } else { // term stores are loaded
        let termStore: SPC.Taxonomy.ITermStore = null;

        if (this._loadedTermStores) {
          for (let i: number = 0, len: number = this._loadedTermStores.get_count(); i < len; i++) {
            if (this._loadedTermStores.get_item(i).get_id().toString() === termStoreId) {
              termStore = this._loadedTermStores.get_item(i);
              break;
            }
          }
        }

        resolve(termStore);
      }
    });
  }

  /**
   * Retrieves term group by its id and parent term store
   * @param termStore: parent term store
   * @param termGroupId: the identifier of the group to retrieve
   */
  private getTermGroupById(termStore: SPC.Taxonomy.ITermStore, termGroupId: string): Promise<SPC.Taxonomy.ITermGroup> {
    return new Promise<SPC.Taxonomy.ITermGroup>((resolve) => {
      if (!termStore || !termGroupId) {
        resolve(null);
        return;
      }

      let result: SPC.Taxonomy.ITermGroup;
      //
      // if Groups property is not loaded get_count will throw an error that will be handled to retrieve groups
      //
      try {
        const groups: SPC.Taxonomy.ITermGroupCollection = termStore.get_groups();
        const groupsCount: number = groups.get_count();
        const groupIdUpper: string = termGroupId.toUpperCase();

        for (let i: number = 0; i < groupsCount; i++) {
          const currGroup: SPC.Taxonomy.ITermGroup = groups.get_item(i);
          if (currGroup.get_id().toString().toUpperCase() === groupIdUpper) {
            result = currGroup;
            break;
          }
        }

        if (!result) { // throwing an exception to try to load the group from server again
          throw new Error();
        }

        resolve(result);
      } catch (ex) { // retrieving the groups from server
        result = termStore.getGroup(termGroupId);
        this.clientContext.load(result);
        this.clientContext.executeQueryAsync(() => {
          resolve(result);
        }, () => {
          resolve(null);
        });
      }
      // tslint:disable-next-line:no-empty
      finally { }
    });
  }

  /**
   * Retrieves term set by its id, parent group and parent store
   * @param termStore: parent term store
   * @param termGroup: parent term group
   * @param termSetId: the identifier of the term set to retrieve
   */
  private getTermSetById(termStore: SPC.Taxonomy.ITermStore, termGroup: SPC.Taxonomy.ITermGroup,
    termSetId: string): Promise<SPC.Taxonomy.ITermSet> {
    return new Promise<SPC.Taxonomy.ITermSet>((resolve) => {
      if (!termGroup || !termSetId) {
        resolve(null);
        return;
      }

      let result: SPC.Taxonomy.ITermSet;
      //
      // if termSets property is not loaded get_count will throw an error that will be handled to retrieve term sets
      //
      try {
        const termSets: SPC.Taxonomy.ITermSetCollection = termGroup.get_termSets();
        const setsCount: number = termSets.get_count();
        const setIdUpper: string = termSetId.toUpperCase();

        for (let i: number = 0; i < setsCount; i++) {
          const currSet: SPC.Taxonomy.ITermSet = termSets.get_item(i);
          if (currSet.get_id().toString().toUpperCase() === setIdUpper) {
            result = currSet;
            break;
          }
        }

        if (!result) { // throwing an exception to try to load the term set from server again
          throw new Error();
        }

        resolve(result);
      } catch (ex) {
        result = termStore.getTermSet(termSetId);
        this.clientContext.load(result);
        this.clientContext.executeQueryAsync(() => {
          resolve(result);
        }, () => {
          resolve(null);
        });
      }
      // tslint:disable-next-line:no-empty
      finally { }
    });
  }

  /**
   * Builds terms' hierarchy and also caches flat list of terms
   * @param terms: SPC.Taxonomy.ITermCollection object
   * @param termSetId: the indetifier of parent term set
   */
  private buildTermsHierarchy(terms: SPC.Taxonomy.ITermCollection, termSetId: string): ITerms {
    if (!terms) {
      return {};
    }

    const tree: ITerms = {};
    const flat: ITerms = {};

    //
    // iterating through terms to collect flat list and create ITermWithTerms instances
    //
    for (let i: number = 0, len: number = terms.get_count(); i < len; i++) {
      const term: SPC.Taxonomy.ITerm = terms.get_item(i);
      // creating instance
      const termEntity: ITermWithTerms = {
        id: term.get_id().toString(),
        name: term.get_name(),
        description: term.get_description(),
        labels: [],
        termsCount: term.get_termsCount(),
        isRoot: term.get_isRoot(),
        path: term.get_pathOfTerm().split(";"),
        fullPath: term.get_pathOfTerm(),
        termSetId: termSetId
      };

      //
      // settings labels
      //
      const labels: SPC.Taxonomy.ILabelCollection = term.get_labels();
      for (let lblIdx: number = 0, lblLen: number = labels.get_count(); lblIdx < lblLen; lblIdx++) {
        const lbl: SPC.Taxonomy.ILabel = labels.get_item(lblIdx);
        termEntity.labels.push({
          isDefaultForLanguage: lbl.get_isDefaultForLanguage(),
          value: lbl.get_value(),
          language: lbl.get_language()
        });
      }

      // if term is root we need to add it to the tree
      if (termEntity.isRoot) {
        tree[termEntity.id] = termEntity;
      }

      // adding term entity to flat list
      flat[termEntity.id] = termEntity;
    }

    const keys: string[] = Object.keys(flat);
    //
    // iterating through flat list of terms to build the tree structure
    //
    for (let keyIdx: number = 0, keysLength: number = keys.length; keyIdx < keysLength; keyIdx++) {
      const key: string = keys[keyIdx];
      const currentTerm: ITermWithTerms = flat[key];

      // skipping root items
      if (currentTerm.isRoot) { continue; }

      // getting parent term name
      const termParentName: string = currentTerm.path[currentTerm.path.length - 2];

      //
      // second iteration to get parent term in flat list
      //
      for (let keySecondIndex: number = 0; keySecondIndex < keysLength; keySecondIndex++) {
        const secondTerm: ITermWithTerms = flat[keys[keySecondIndex]];
        if (secondTerm.name === termParentName && secondTerm.path.length === currentTerm.path.length - 1 &&
          currentTerm.fullPath.indexOf(secondTerm.fullPath) === 0) {
          if (!secondTerm.terms) {
            secondTerm.terms = {};
          }
          secondTerm.terms[currentTerm.id] = currentTerm;
        }
      }
    }

    this._loadedTermsFlat.push(flat);

    return tree;
  }

  /**
   * Converts ITerms object to collection of ITerm objects
   * @param terms: ITerms object
   */
  private getTermEntities(terms: ITerms): ITerm[] {
    const termsKeys: string[] = Object.keys(terms);
    const termEntities: ITerm[] = [];
    for (let keyIdx: number = 0, keysLength: number = termsKeys.length; keyIdx < keysLength; keyIdx++) {
      termEntities.push(terms[termsKeys[keyIdx]]);
    }

    return termEntities;
  }
}

