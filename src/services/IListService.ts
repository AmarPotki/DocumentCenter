import { IList, IPage } from "../domains";
import { IServiceBase } from "./IServiceBase";
export interface IListService extends IServiceBase {
  /**
   * API to get all pages of the pages list
   */
  getPages: (listName: string) => Promise<IPage[]>;
  /**
   * API to get all lists of the current web
   */
  getLists: () => Promise<IList[]>;
}