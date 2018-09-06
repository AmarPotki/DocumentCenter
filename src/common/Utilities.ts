import { IUtilities } from "./IUtilities";
const iconPdfLarge: any = require("../assets/icon_PDF_Large.png");
const iconWordLarge: any = require("../assets/icon_word_Large.png");
const iconExcelLarge: any = require("../assets/icon_excel_Large.png");
const iconPngLarge: any = require("../assets/icon_png_Large.png");
const iconPptLarge: any = require("../assets/icon_PowerPoint_Large.png");
const iconDocumentLarge: any = require("../assets/icon_Document_Large.png");
const iconPdfSmall: any = require("../assets/icon_PDF_Small.png");
const iconWordSmall: any = require("../assets/icon_word_Small.png");
const iconExcelSmall: any = require("../assets/icon_excel_Small.png");
const iconPngSmall: any = require("../assets/icon_png_Small.png");
const iconPptSmall: any = require("../assets/icon_PowerPoint_Small.png");
const iconDocumentSmall: any = require("../assets/icon_Document_Small.png");
const iconMicrophoneSmall: any = require("../assets/icon_Microphone_Small.png");
const iconMicrophoneSmallRed: any = require("../assets/icon_Microphone_Small_Red.png");
const iconMagnifierSmall: any = require("../assets/icon_Magnifier_Small.png");
const iconLargFolder: any = require("../assets/icon_Folder_Large.png");
const iconDropDown: any = require("../assets/icon-dropdown.png");
const iconTranslate: any = require("../assets/Action-Translate.png");
const iconShare: any = require("../assets/Action-Share.png");
const iconBookmark: any = require("../assets/Action-Bookmark.png");

export class Utilities implements IUtilities {
    /**
     * API to get date from ISOString formated date
     */
    public getDateFromISOString(isoDate: string): string {
        // define a new date object based on iso date format
        var d: Date = new Date(isoDate);
        var date: string = "";
        // join the year,month and day value of the date as MM/DD/YYYY
        date += d.getDate() + "/";
        var month: string = (d.getMonth() + 1).toString();
        // add 0 to the month value if it's less than 10 ( by default it's not included)
        if (month.length === 1) {
            month = "0" + month;
        }
        date += month + "/" + d.getFullYear();
        // return the date
        return date;
    }
    /**
     * API to get document icon by its extension
     */
    public getDocumentIcon(extension: string, large: boolean): any {
        // find the relevant icon for that extension
        switch (extension) {
            case "pdf":
                return large ? iconPdfLarge : iconPdfSmall;
            case "docx":
            case "doc":
                return large ? iconWordLarge : iconWordSmall;
            case "xlsx":
            case "xls":
                return large ? iconExcelLarge : iconExcelSmall;
            case "png":
                return large ? iconPngLarge : iconPngSmall;
            case "Powerpoint":
                return large ? iconPptLarge : iconPptSmall;
            default:
                return large ? iconDocumentLarge : iconDocumentSmall;
        }
    }
    public getIcon(name: string, large: boolean): any {
        switch (name) {
            case "Microphon":
                return large ? "" : iconMicrophoneSmall;
            case "MicrophonRed":
                return large ? "" : iconMicrophoneSmallRed;
            case "Magnifier":
                return large ? "" : iconMagnifierSmall;
            case "Translate":
                return large ? "" : iconTranslate;
            case "Share":
                return large ? "" : iconShare;
            case "Bookmark":
                return large ? "" : iconBookmark;
            case "DropDown":
                return large ? iconDropDown : iconDropDown;
            case "Folder":
                return large ? iconLargFolder : "";
            default:
                return large ? iconDocumentLarge : iconDocumentSmall;
        }
    }
    public stripHTMLFromText(input: string): string {
        return input.replace(/<[^>]*>/g, '');
    }
}