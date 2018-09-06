import {
    Environment,
    EnvironmentType
} from "@microsoft/sp-core-library";
import { IServiceBase } from "./IServiceBase";
import { IServiceFactory } from "./IServiceFactory";
import { IMockServiceBase } from "./IMockServiceBase";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
/**
 * Factory object to create service helper based on current EnvironmentType
 */
export class ServiceFactory implements IServiceFactory {
    /**
     * API to create service helper
     * @context: web part context
     */
    public GetService<T extends IServiceBase>(context: IWebPartContext,
        service: new (c: IWebPartContext) => T, mockService: new () => T): T {
        // local environment
        if (Environment.type === EnvironmentType.Local) {
            return new mockService();
        } else {
            return new service(context);
        }
    }
}