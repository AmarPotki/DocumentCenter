import { IList,IPage } from '../domains';
import {IListService} from '../services';

export class MockListService implements IListService {
  private lists: IList[];
  private items: any;

  constructor() {
    /**
   * MOCK lists data
   */
    this.lists = [      
      {
        Id: "Pages",
        Title: 'Pages',
        ServerRelativeUrl:""
      },
      {
        Id:"Documents",
        Title:"Documents",
        ServerRelativeUrl:""
      },
      {
        Id:"Tasks",
        Title:"Tasks",
        ServerRelativeUrl:""
      }
    ];
    /**
   * MOCK items
   */
    this.items = {
        Pages: [
          {
            Id: '1',
            Name: 'Home.aspx'
          },
          {
            Id: '2',
            Name: 'Default.aspx'
          },
          {
            Id: '3',
            Name: 'SearchResults.aspx'
          }
        ]
      };
  }
  /**
 * MOCK data helper. Gets pages from hardcoded values
 */
  public getLists():Promise<IList[]>{
    return new Promise<IList[]>(resolve => {
      setTimeout(() => {
        resolve(this.lists);
      }, 1000);
    }); 
  }
/**
 * MOCK data helper. Gets lists from hardcoded values
 */
  public getPages(listName: string): Promise<IPage[]> {
    return new Promise<IPage[]>(resolve => {
      setTimeout(() => {
        resolve(this.items[listName]);
      }, 1000);
    });
  }
}