/**
 * @interface
 * Defines a SharePoint UpcomingBirthday list items
 */
export interface IUpcomingBirthdayList{
    value:IUpcomingBirthdayItem[];
    }

    /**
 * @interface
 * Defines a SharePoint UpcomingBirthday list item
 */
    export interface IUpcomingBirthdayItem{
    Title:string;
    Description:string;
    picture:string;
    Created:string;
    Currentdate:string;
    ID:string;
    }