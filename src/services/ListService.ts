import { IPage, IResponseFile, IList } from "../domains";
import { IListService } from "./IListService";
import { Web, ListEnsureResult } from "sp-pnp-js";
import * as pnp from "sp-pnp-js";
import { IWebPartContext } from "@microsoft/sp-webpart-base";

export class ListService implements IListService {

  constructor(private context: IWebPartContext) {
    /**
     * Setup pnp to use current context
     */
    pnp.setup({
      spfxContext: this.context
    });
  }
  /**
   * API to get all lists of the current web
   */
  public async getLists(): Promise<IList[]> {
    let response: any = await pnp.sp.web.lists.expand("RootFolder").get();
    let lists: IList[] = response.map((item: any) => {
      var relativeUrl: string = item.RootFolder.ServerRelativeUrl;
      return {
        Title: item.Title,
        Id: item.Id,
        ServerRelativeUrl: relativeUrl.substr(relativeUrl.lastIndexOf("/") + 1, relativeUrl.length)
      };
    });
    return Promise.resolve(lists);
  }
  /**
   * API to get all pages of the pages list
   */
  public getPages(listName: string): Promise<IPage[]> {
    // const web: Web = new Web(this.context.pageContext.web.absoluteUrl);
    return new Promise<IPage[]>((resolve: (results: IPage[]) => void, reject: (error: any) => void): void => {
      pnp.sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl + "/SitePages")
        .files.get().then((response: IResponseFile[]) => {
          let pages: IPage[] = [];
          pages = response.map((item: IResponseFile) => {
            return {
              Name: item.Name,
              Id: item.Id
            };
          });
          resolve(pages);
        });
    });
  }
}