export interface IOffice365Group {
    // Microsoft Graph has more group properties.
    displayName: string;
    mail: string;
    description: string;
}

export interface IGraphApiService {
    readGroup(): void;
    sendMail(mailAddress: string,subject:string, body: string): void;
     GetRecentlyViewed(): void;
}