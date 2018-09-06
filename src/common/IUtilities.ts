export interface IUtilities {
    /**
     * API to get date from ISOString formated date
     */

    getDateFromISOString: (isoDate: string) => string;
    /**
     * API to get document icon by its extension
     */
    getDocumentIcon: (extension: string, large: boolean) => any;
    getIcon: (name: string, large: boolean) => any;
    stripHTMLFromText(input: string): string;
}