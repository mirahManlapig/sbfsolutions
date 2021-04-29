/**
 * @interface
 * Defines a SharePoint Announcements list items
 */
export interface IAnnouncementsList{
    value:IAnnouncementsItem[];
    }

    /**
 * @interface
 * Defines a SharePoint Announcements list item
 */
    export interface IAnnouncementsItem{
    Title:string;
    Description:string;
    Effectivedate:string;
    Expirydate:string;
    Created:string;
    ID:string;


    
    }