import { IDocumentItem, IContentType, IResponseFile, ICCDocument, IAssociatedDocument, IResponseAssociatedDocument } from '../domains';
import { IDocumentService } from './IDocumentService';
import { Web, ListEnsureResult } from "sp-pnp-js";
import * as pnp from 'sp-pnp-js';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import {IUtilities,Utilities} from '../common';
export class MockDocumentService implements IDocumentService {
 private utilities : IUtilities;
  constructor() {
    this.utilities=new Utilities();
    
  }
  /**
     * API to get Documents from a library
  */
  public async getDocumentsByListName(library:string,top:number,properties:string,order:string,ascending:boolean):Promise<IDocumentItem[]>{
    return new Promise<IDocumentItem[]>(resolve => {
      setTimeout(() => {
        resolve([
          {
            DocIcon:".pdf",
            Title:"Management Of Medical Gases",
            Id:2,
            Icon:this.utilities.getDocumentIcon("pdf",true),
            File:null,
            Created:"15/10/2017"
          }
        ]);
      }, 1000);
    });
  }
  /**
     * API to get document properties by document path
  */
  public async getDocumentByPath(documentPath:string): Promise<ICCDocument> {    
    return new Promise<ICCDocument>((resolve: (results: ICCDocument) => void, reject: (error: any) => void): void => {     
          let document: ICCDocument = <ICCDocument>{};
          document.Id = 1153;
          document.Icon = this.utilities.getDocumentIcon("pdf",true);
          document.Name = "Management Of Medical Gases.pdf";
          document.Title = "Management Of Medical Gases";
          document.BMIDocumentCategory = "Policy";
          document.Created = "15/10/2017";
          document.IssueDate = "15/10/2017";
          document.ContentTypeId = "0x0101006E2D115A075FB0449CC85F298C9EA8010035F0DCA4E032094F99A798C973DB9548";
          document.ServerRelativeUrl = "/sites/bmidocuments/test/Management Of Medical Gases.pdf";
          document.DocumentSetFolder = "/sites/bmidocuments/test";
          document.UniqueId ="{7aa4b5ed-d57e-4e8b-9004-3c91e8f6c01a}";
          document.Path="https://bmihcqa/sites/bmidocuments/documents/Management Of Medical Gases.pdf";
          document.OnlinePath=`https://bmihcqa/sites/bmidocuments/_layouts/15/WopiFrame.aspx?sourcedoc={7aa4b5ed-d57e-4e8b-9004-3c91e8f6c01a}&file=Management Of Medical Gases.pdf&action=default`;
          resolve(document);
    });
  }
  /**
     * API to get document properties by document ID
  */
  public async getDocumentById(documentId: number,documentLibrary:string): Promise<ICCDocument> {    
    return new Promise<ICCDocument>((resolve: (results: ICCDocument) => void, reject: (error: any) => void): void => {     
          let document: ICCDocument = <ICCDocument>{};
          document.Id = 1153;
          document.Icon = this.utilities.getDocumentIcon("pdf",true);
          document.Name = "Management Of Medical Gases.pdf";
          document.Title = "Management Of Medical Gases";
          document.BMIDocumentCategory = "Policy";
          document.Created = "15/10/2017";
          document.ContentTypeId = "0x0101006E2D115A075FB0449CC85F298C9EA8010035F0DCA4E032094F99A798C973DB9548";
          document.ServerRelativeUrl = "/sites/bmidocuments/test/Management Of Medical Gases.pdf";
          document.DocumentSetFolder = "/sites/bmidocuments/test";
          document.UniqueId ="{7aa4b5ed-d57e-4e8b-9004-3c91e8f6c01a}";
          document.Path="https://bmihcqa/sites/bmidocuments/documents/Management Of Medical Gases.pdf";
          document.OnlinePath=`https://bmihcqa/sites/bmidocuments/_layouts/15/WopiFrame.aspx?sourcedoc={7aa4b5ed-d57e-4e8b-9004-3c91e8f6c01a}&file=Management Of Medical Gases.pdf&action=default`;
          resolve(document);
    });
  }
  /**
     * API to get related documents from a library
  */
  public async getRelatedDocumentsByPath(library:string,documentPath:string,top:number,orderby:string,ascending:boolean):Promise<IDocumentItem[]>{
    return new Promise<IDocumentItem[]>(resolve => {
      setTimeout(() => {
        resolve([
          {
            DocIcon:".pdf",
            Title:"Management Of Medical Gases",
            Id:2,
            Icon:this.utilities.getDocumentIcon("pdf",true),
            File:null,
            Created:"15/10/2017"
          },
          {
            DocIcon:".pdf",
            Title:"BMI NURman01 Urgent Care Centres V2.0",
            Id:123,
            Icon:this.utilities.getDocumentIcon("pdf",true),
            File:null,
            Created:"12/10/2017"
          },
          {
            DocIcon:".doc",
            Title:"BMI GOVpol01 - Temp03 Local WI Template Final V1.0",
            Id:414,
            Icon:this.utilities.getDocumentIcon("doc",true),
            File:null,
            Created:"11/09/2017"
          },
          {
            DocIcon:".doc",
            Title:"BMI PATHsop07 - Form01 24 Hour Urinary 5HIAA V1.0",
            Id:2,
            Icon:this.utilities.getDocumentIcon("doc",true),
            File:null,
            Created:"07/10/2017"
          },
          {
            DocIcon:".pdf",
            Title:"BMI NURpol30 Theatre Late Booking V1.0",
            Id:2,
            Icon:this.utilities.getDocumentIcon("pdf",true),
            File:null,
            Created:"02/10/2017"
          }
        ]);
      }, 1000);
    });
  }
   /**
     * API to get Documents of a document set folder
  */
  public async getDocumentSetDocuments(folderName:string):Promise<any[]>{
    const response = await pnp.sp.web.getFolderByServerRelativeUrl(folderName).files
      .expand("ListItemAllFields")
      .expand("ListItemAllFields,ListItemAllFields/ContentType")
      .orderBy("ListItemAllFields/DocumentOrder")
      .get();
    return Promise.resolve(response);
  }
  /**
     * API to get taxonomy label from TaxonomyHiddenList
     * we are using this API because when you make a query over a item with
       single value taxonomy field , it doesn't return the label of that taxonomy field
  */
  public async getAssociatedDocuments(folderName: string): Promise<IAssociatedDocument[]> {
    return new Promise<IAssociatedDocument[]>(resolve => {
      setTimeout(() => {
        resolve([
          {
            ContentTypeId:"0x0101006E2D115A075FB0449CC85F298C9EA80102008E9F2CFB375F4D42821059976091DE3B",
            Created:"17/10/2017",
            Icon:this.utilities.getDocumentIcon("pdf",true),
            Id:1156,
            Name:"sets out BMI Healthcare’s principles",
            Title:"sets out BMI Healthcare’s principles",
            ServerRelativeUrl:"/sites/bmidocuments/documents/sets out BMI Healthcare’s principles.pdf",
            Path:"sets out BMI Healthcare’s principles",
            OnlinePath:"sets out BMI Healthcare’s principles",
            UniqueId:"{7aa4b5ed-d57e-4e8b-9004-3c91e8f6c01b}",
            FileType:"",
            TimeLastModified:"",
          },
          {
            ContentTypeId:"0x0101006E2D115A075FB0449CC85F298C9EA80102008E9F2CFB375F4D42821059976091DE3B",
            Created:"17/10/2017",
            Icon:this.utilities.getDocumentIcon("docx",true),
            Id:1157,
            Name:"BMI IMpol15 Retention of Records V3.0",
            Title:"BMI IMpol15 Retention of Records V3.0",
            ServerRelativeUrl:"/sites/bmidocuments/documents/BMI IMpol15 Retention of Records V3.0.docx",
            Path:"BMI IMpol15 Retention of Records V3.0",
            OnlinePath:"BMI IMpol15 Retention of Records V3.0",
            UniqueId:"{7aa4b5ed-d57e-4e8b-9004-3c91e8f6c01c}",
            FileType:"",
            TimeLastModified:"",
          }
        ]);
      }, 1000);
    });
  }

  
}