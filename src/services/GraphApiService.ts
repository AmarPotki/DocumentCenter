import { GraphHttpClient, HttpClientResponse, IGraphHttpClientOptions, GraphHttpClientResponse } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IOffice365Group, IGraphApiService } from './IGraphApiService';
export class GraphApiService implements IGraphApiService {
    private context: IWebPartContext;
    constructor(pageContext: IWebPartContext) {
        this.context = pageContext;
    }
    public readGroup(): void {
        this.context.graphHttpClient.get(`v1.0/groups?$orderby=displayName`, GraphHttpClient.configurations.v1).then((response: HttpClientResponse) => {
            if (response.ok) {
                return response.json();
            } else {
                console.warn(response.statusText);
            }
        }).then((result: any) => {
            // Transfer result values to the group variable
            console.log("ok");
            console.log(result.value);
            console.log(this.renderTable(result.value));
        });
    }
    protected renderTable(items: IOffice365Group[]): string {
        let html: string = '';
        if (items.length <= 0) {
            html = `<p>There are no groups to list...</p>`;
        }
        else {
            html += `
          <table><tr>
            <th>Display Name</th>
            <th>Mail</th>
            <th>Description</th></tr>`;
            items.forEach((item: IOffice365Group) => {
                html += `
              <tr>
                  <td>${item.displayName}</td>
                  <td>${item.mail}</td>
                  <td>${item.description}</td>
              </tr>`;
            });
            html += `</table>`;
        }
        return html;
        //  const tableContainer: Element = this.domElement.querySelector('#spTableContainer');
        //  tableContainer.innerHTML = html;
        // return;
    }

    public sendMail(mailAddress: string, subject: string, body: string): void {
        const requestHeaders: Headers = new Headers();
        requestHeaders.append('Content-type', 'application/json');
        requestHeaders.append('Content-length', '512');
        let mail: IGraphHttpClientOptions = {
            body: {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "Text",
                        "content": body
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": mailAddress
                            }
                        }
                    ],
                    "ccRecipients": [
                        {
                            "emailAddress": {
                                "address": "ashkan.shirian@cielocosta.com"
                            }
                        }
                    ]
                },
                "saveToSentItems": "false"
            },
            headers: requestHeaders
        };

        this.context.graphHttpClient.post("v1.0/me/sendMail", GraphHttpClient.configurations.v1, mail).then((response: GraphHttpClientResponse) => {
            if (response.ok) {
                return response.json();
            } else {
                console.warn(response.statusText);
            }
        }).then((result: any) => {
            // Transfer result values to the group variable
            console.log("ok");
            console.log(result.value);
        });
    }

    public GetRecentlyViewed(): void {
        this.context.graphHttpClient.get("https://graph.microsoft.com/beta/me/insights/used?$filter=ResourceVisualization/Type eq 'Excel'",
         GraphHttpClient.configurations.v1).then((response: GraphHttpClientResponse) => {
            if (response.ok) {
                console.log(response);
                return response.json();
            } else {
                console.warn(response.statusText);
            }
        }).then((result: any) => {
            // Transfer result values to the group variable
            console.log("ok");
            console.log(result.value);
        });
    }
}