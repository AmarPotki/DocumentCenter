import{IServiceBase} from './IServiceBase';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
export interface IServiceFactory{
    /**
   * API to create service helper
   * @context: web part context
   */
    GetService<T extends IServiceBase>(context:IWebPartContext,service: new (c:IWebPartContext) => T,mockService: new () => T):T;
}