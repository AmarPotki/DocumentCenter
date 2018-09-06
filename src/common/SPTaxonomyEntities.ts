/**
 * Base interface for Taxonomy objects
 */
export interface ITermBase {
  id: string;
  name: string;
}

/**
 * Term Store interface
 */
export interface ITermStore extends ITermBase {
}

/**
 * Term Group interface
 */
export interface ITermGroup extends ITermBase {
  description: string;
  termStoreId: string;
}

/**
 * Term Set Interface
 */
export interface ITermSet extends ITermBase {
  description?: string;
  termGroupId?: string;
  termStoreId?: string;
}

/**
 * Term interface
 */
export interface ITerm extends ITermBase {
  description: string;
  isRoot: boolean;
  labels: ILabel[];
  termsCount: number;
  termSetId: string;
}

/**
 * Term Label interface
 */
export interface ILabel {
  isDefaultForLanguage: boolean;
  value: string;
  language: string;
}