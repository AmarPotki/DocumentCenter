import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { PrincipalType, IPropertyFieldGroupOrPerson } from "./../controls/peoplePicker/IPropertyFieldPeoplePicker";
import { ISPPeopleSearchService } from "./ISPPeopleSearchService";
import SPPeoplePickerMockHttpClient from "./SPPeopleSearchMockService";

/**
 * Service implementation to search people in SharePoint
 */
export default class SPPeopleSearchService implements ISPPeopleSearchService {
  private context: IWebPartContext;

  /**
   * Service constructor
   */
  constructor(pageContext: IWebPartContext) {
    this.context = pageContext;
  }

  /**
   * Search people from the SharePoint People database
   */
  public searchPeople(query: string, principalType: PrincipalType[]): Promise<Array<IPropertyFieldGroupOrPerson>> {
    if (Environment.type === EnvironmentType.Local) {
      // if the running environment is local, load the data from the mock
      return this.searchPeopleFromMock(query);
    } else {
      // if the running env is SharePoint, loads from the peoplepicker web service
      // tslint:disable-next-line:max-line-length
      const userRequestUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
      // tslint:disable-next-line:typedef
      const data = {
        "queryParams": {
          "AllowEmailAddresses": true,
          "AllowMultipleEntities": false,
          "AllUrlZones": false,
          "MaximumEntitySuggestions": 20,
          "PrincipalSource": 15,
          // principalType controls the type of entities that are returned in the results.
          // choices are All - 15, Distribution List - 2 , Security Groups - 4, SharePoint Groups - 8, User - 1.
          // these values can be combined (example: 13 is security + SP groups + users)
          "PrincipalType": !!principalType && principalType.length > 0 ? principalType.reduce((a, b) => a + b, 0) : 1,
          "QueryString": query
        }
      };

      const requestHeaders: Headers = new Headers();
      requestHeaders.append("accept", "application/json");
      requestHeaders.append("content-type", "application/xml");
      let httpPostOptions: ISPHttpClientOptions = {
        headers: requestHeaders,
        body: JSON.stringify(data)
      };
      // old code
      // let httpPostOptions: ISPHttpClientOptions = {
      //   headers: {
      //     'accept': 'application/json',
      //     'content-type': 'application/json'
      //   },
      //   body: JSON.stringify(data)
      // };

      // do the call against the People REST API endpoint
      // tslint:disable-next-line:max-line-length
      return this.context.spHttpClient.post(userRequestUrl, SPHttpClient.configurations.v1, httpPostOptions).then((searchResponse: SPHttpClientResponse) => {
        return searchResponse.json().then((usersResponse: any) => {
          const res: IPropertyFieldGroupOrPerson[] = [];
          const values: any = JSON.parse(usersResponse.value);
          values.map(element => {
            switch (element.EntityType) {
              case "User":
                const groupOrPerson: IPropertyFieldGroupOrPerson = { fullName: element.DisplayText, login: element.Description };
                groupOrPerson.email = element.EntityData.Email;
                groupOrPerson.jobTitle = element.EntityData.Title;
                groupOrPerson.initials = this.getFullNameInitials(groupOrPerson.fullName);
                groupOrPerson.imageUrl = this.getUserPhotoUrl(groupOrPerson.email, this.context.pageContext.web.absoluteUrl);
                res.push(groupOrPerson);
                break;
              case "SecGroup":
                const group: IPropertyFieldGroupOrPerson = {
                  fullName: element.DisplayText,
                  login: element.ProviderName,
                  id: element.Key,
                  description: element.Description
                };
                res.push(group);
                break;
              default:
                const persona: IPropertyFieldGroupOrPerson = {
                  fullName: element.DisplayText,
                  login: element.EntityData.AccountName,
                  id: element.EntityData.SPGroupID,
                  description: element.Description
                };
                res.push(persona);
                break;
            }
          });
          return res;
        });
      });
    }
  }

  /**
   * Generates Initials from a full name
   */
  private getFullNameInitials(fullName: string): string {
    if (fullName === null) {
      return fullName;
    }

    const words: string[] = fullName.split(" ");
    if (words.length === 0) {
      return "";
    } else if (words.length === 1) {
      return words[0].charAt(0);
    } else {
      return (words[0].charAt(0) + words[1].charAt(0));
    }
  }

  /**
   * Gets the user photo url
   */
  private getUserPhotoUrl(userEmail: string, siteUrl: string): string {
    return `${siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${userEmail}`;
  }


  /**
   * Returns fake people results for the Mock mode
   */
  private searchPeopleFromMock(query: string): Promise<Array<IPropertyFieldGroupOrPerson>> {
    return SPPeoplePickerMockHttpClient.searchPeople(this.context.pageContext.web.absoluteUrl).then(() => {
      const results: IPropertyFieldGroupOrPerson[] = [
        { fullName: "Katie Jordan", initials: "KJ", jobTitle: "VIP Marketing", email: "katiej@contoso.com", login: "katiej@contoso.com" },
        { fullName: "Gareth Fort", initials: "GF", jobTitle: "Sales Lead", email: "garethf@contoso.com", login: "garethf@contoso.com" },
        { fullName: "Sara Davis", initials: "SD", jobTitle: "Assistant", email: "sarad@contoso.com", login: "sarad@contoso.com" },
        { fullName: "John Doe", initials: "JD", jobTitle: "Developer", email: "johnd@contoso.com", login: "johnd@contoso.com" }
      ];
      return results;
    }) as Promise<Array<IPropertyFieldGroupOrPerson>>;
  }
}
