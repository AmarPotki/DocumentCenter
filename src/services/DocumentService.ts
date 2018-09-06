import {
  ITaxonomyItem, IContentType, IResponseFile, IDocumentItem,
  ICCDocument, IAssociatedDocument, IResponseAssociatedDocument
} from "../domains";
import { IDocumentService } from "./IDocumentService";
import { Web, ListEnsureResult, Dictionary } from "sp-pnp-js";
import * as pnp from "sp-pnp-js";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { IUtilities, Utilities } from "../common";
import { QueryableInstance } from "sp-pnp-js/lib/sharepoint/queryable";
export class DocumentService implements IDocumentService {
  private utilities: IUtilities;
  constructor(private context: IWebPartContext) {
    /**
     * Setup pnp to use current context
     */
    this.utilities = new Utilities();
    pnp.setup({
      spfxContext: this.context
    });
  }
  /**
   * API to get Documents from a library
   */
  public async getDocumentsByListName(library: string, top: number, properties: string,
    order: string, ascending: boolean): Promise<IDocumentItem[]> {
    let response: any = await pnp.sp.web.lists.getByTitle(library).items
      .expand("File,ContentType").filter("ContentType eq 'CCDocument' or ContentType eq 'CCAssociatedDocument'")
      .top(top).select(properties).orderBy(order, ascending).get();

    let documents: IDocumentItem[] = [];
    documents = response.map((item) => {
      return {
        Title: item.File.Name.substr(0, item.File.Name.lastIndexOf(".")),
        Id: item.Id,
        Icon: this.utilities.getDocumentIcon(item.File.Name.split(".").pop(), false),
        File: item.File,
        Created: item.Created
      };
    });
    return Promise.resolve(documents);
  }

  /**
   * API to get document properties by document path
   */
  public async getDocumentByPath(documentPath: string, webUrl: string): Promise<ICCDocument> {
    let document: ICCDocument = <ICCDocument>{};
    let tenantUrl: string = window.location.protocol + "//" + window.location.host;
    let web = new Web(webUrl);
    const response: pnp.Item = await web.getFileByServerRelativeUrl(documentPath).getItem();
    const fieldValues: QueryableInstance = await response.fieldValuesAsHTML.get();
    let fields: Dictionary<string> = new Dictionary<string>();
    for (var k in fieldValues) {
      fields.add(k, fieldValues[k]);
    }
    document.fields = fields;
    const documentResponse = await response.expand("File,ContentType").get();

    let extension: string = documentResponse.File.Name.split(".").pop();
    extension = extension.toLowerCase();
    document.Id = documentResponse.Id;
    document.Icon = this.utilities.getDocumentIcon(extension, true);
    document.Name = documentResponse.File.Name;
    document.Title = documentResponse.File.Name.substr(0, documentResponse.File.Name.lastIndexOf("."));
    document.Created = this.utilities.getDateFromISOString(documentResponse.Created);
    document.IssueDate = documentResponse.IssueDate != null ? this.utilities.getDateFromISOString(documentResponse.IssueDate) : "";
    document.ContentTypeName = documentResponse.ContentType.Name;
    document.DocumentSetFolder = documentResponse.File.ServerRelativeUrl.substr(0, documentResponse.File.ServerRelativeUrl.lastIndexOf("/"));
    document.ServerRelativeUrl = documentResponse.File.ServerRelativeUrl;
    document.UniqueId = `{${documentResponse.File.UniqueId}}`;
    document.Path = tenantUrl + documentResponse.File.ServerRelativeUrl;
    document.OnlinePath = extension !== "pdf" ? `${this.context.pageContext.web.absoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${documentResponse.File.UniqueId}&file=${documentResponse.File.Name}&action=default` : document.Path;

    // let bmiCategory: ITaxonomyItem = <ITaxonomyItem>{};
    // if (documentResponse.BMI_x0020_Document_x0020_Category != null) {
    //   let bmiTaxonomy = await this.getTaxonomyTitle(documentResponse.BMI_x0020_Document_x0020_Category.TermGuid);
    //   var title = bmiTaxonomy.Title === "" ? "Unknown" : bmiTaxonomy.Title;
    //   document.BMIDocumentCategory = title;
    // }
    // else {
    //   document.BMIDocumentCategory = "Unknown";
    // }
    return Promise.resolve(document);
  }
  /**
     * API to get document properties by document ID
  */
  public async getDocumentById(documentId: number): Promise<ICCDocument> {
    // const web: Web = new Web(this.context.pageContext.web.absoluteUrl);
    let document: ICCDocument = <ICCDocument>{};
    const documentResponse = await pnp.sp.web.lists
      .getByTitle("Documents").items
      .getById(documentId)
      .expand("File")
      .get();
    document.Id = documentResponse.Id;
    document.Icon = this.utilities.getDocumentIcon(documentResponse.File.Name.split(".").pop(), true);
    document.Name = documentResponse.File.Name;
    document.Title = documentResponse.File.Name.substr(0, documentResponse.File.Name.lastIndexOf("."));
    //document.BMIDocumentCategory = bmiCategory.Title;
    document.Created = this.utilities.getDateFromISOString(documentResponse.Created);
    document.ContentTypeId = documentResponse.ContentTypeId;
    document.DocumentSetFolder = documentResponse.File.ServerRelativeUrl.substr(0, documentResponse.File.ServerRelativeUrl.lastIndexOf("/"));
    document.ServerRelativeUrl = documentResponse.File.ServerRelativeUrl;
    // document.PrimaryDocument = item.PrimaryDocument;
    document.UniqueId = `{${documentResponse.File.UniqueId}}`;
    document.Path = this.context.pageContext.web.absoluteUrl + documentResponse.File.ServerRelativeUrl;
    document.OnlinePath = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc={${documentResponse.File.UniqueId}}&file=${documentResponse.File.Name}&action=default`;


    let bmiCategory: ITaxonomyItem = <ITaxonomyItem>{};
    if (documentResponse.BMI_x0020_Document_x0020_Category != null) {
      let bmiTaxonomy = await this.getTaxonomyTitle(documentResponse.BMI_x0020_Document_x0020_Category.TermGuid);
      document.BMIDocumentCategory = bmiTaxonomy.Title;
    }
    else
      document.BMIDocumentCategory = "Unknown";
    return Promise.resolve(document);
  }
  /**
     * API to get associated documents from a folder
  */
  public async getAssociatedDocuments(folderName: string, webUrl: string): Promise<IAssociatedDocument[]> {
    let associatedDocuments: IAssociatedDocument[] = [];
    let web = new Web(webUrl);
    const response = await web.getFolderByServerRelativeUrl(folderName).files
      .expand("ListItemAllFields")
      .expand("ListItemAllFields,ListItemAllFields/ContentType")
      //.filter("ListItemAllFields/ContentType/Name eq 'CCAssociatedDocument'")
      // .orderBy("ListItemAllFields/DocumentOrder")
      .get();
    associatedDocuments = response.map((file: IResponseAssociatedDocument) => {
      return {
        Id: file.ListItemAllFields.Id,
        Name: file.Name,
        Icon: this.utilities.getDocumentIcon(file.Name.split(".").pop(), true),
        Title: file.Title === "" ? file.Name.substr(0, file.Name.lastIndexOf(".")) : file.Title,
        FileType: file.Name.split(".").pop(),
        Created: this.utilities.getDateFromISOString(file.ListItemAllFields.Created),
        ContentTypeId: file.ListItemAllFields.ContentTypeId,
        ServerRelativeUrl: file.ServerRelativeUrl,
        primaryDocument: file.ListItemAllFields,
        TimeLastModified: file.TimeLastModified,
        UniqueId: `{${file.UniqueId}}`,
        Path: this.context.pageContext.web.absoluteUrl + file.ServerRelativeUrl,
        OnlinePath: `${this.context.pageContext.web.absoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc={${file.UniqueId}}&file=${file.Name}&action=default`
      };
    });
    return Promise.resolve(associatedDocuments);
  }
  /**
    * API to get Documents of a document set folder
 */
  public async getDocumentSetDocuments(folderName: string): Promise<any[]> {

    const response = await pnp.sp.web.getFolderByServerRelativeUrl(folderName).files
      .expand("ListItemAllFields")
      .expand("ListItemAllFields,ListItemAllFields/ContentType")
      // .orderBy("ListItemAllFields/DocumentOrder")
      .get();
    return Promise.resolve(response);
  }
  /**
     * API to get related documents from a library
  */
  public async getRelatedDocumentsByPath(library: string, documentPath: string, top: number, orderby: string, ascending: boolean): Promise<IDocumentItem[]> {
    const documentResponse = await pnp.sp.web.getFileByServerRelativeUrl(documentPath).getItem();
    const documentItemResponse = await documentResponse.select("Id,BMI_x0020_Function").get();
    let documentId: number = documentItemResponse.Id;
    let categoryId: number = documentItemResponse.BMI_x0020_Function.WssId;
    let relatedDocuments: IDocumentItem[] = [];
    let query: string = `<View><Query><OrderBy><FieldRef Name="${orderby}" Ascending="${ascending}"/></OrderBy><Where><And><Eq><FieldRef LookupId="True" Name="BMI_x0020_Function"/><Value Type="Integer">${categoryId}</Value></Eq><Neq><FieldRef Name="ID"/><Value Type="Number">${documentId}</Value></Neq></And></Where></Query><RowLimit>${top}</RowLimit></View>`;
    const relatedDocsResponse = await pnp.sp.web.lists.getByTitle(library).getItemsByCAMLQuery({ ViewXml: query }, "File");
    relatedDocuments = relatedDocsResponse.map((item) => {
      return {
        Title: item.File.Name.substr(0, item.File.Name.lastIndexOf(".")),
        Id: item.Id,
        Icon: this.utilities.getDocumentIcon(item.File.Name.split(".").pop(), false),
        File: item.File,
        Created: item.Created
      };
    });
    return Promise.resolve(relatedDocuments);
  }
  /**
     * API to get taxonomy label from TaxonomyHiddenList
     * we are using this API because when you make a query over a item with
       single value taxonomy field , it doesn't return the label of that taxonomy field
  */
  public async getTaxonomyTitle(taxId: string): Promise<ITaxonomyItem> {
    const response = await pnp.sp.web.lists.getByTitle("TaxonomyHiddenList").items.select("Id,Title").filter("IdForTerm eq '" + taxId + "'").top(1).get();
    let taxonomy: ITaxonomyItem = <ITaxonomyItem>{ Title: response[0].Title, Id: response[0].Id };
    return Promise.resolve(taxonomy);
  }

}