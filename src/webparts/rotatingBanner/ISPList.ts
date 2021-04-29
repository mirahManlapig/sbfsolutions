/**
 * @interface
 * Defines a SharePoint RotatingBanner list items
 */
export interface IRotatingBannerList{
    value:IRotatingBannerItem[];
    }

    /**
 * @interface
 * Defines a SharePoint RotatingBanner list item
 */
    export interface IRotatingBannerItem{
    Title:string;
    BannerDescription:string;
    PreviewImageurl:string;
    FileLeafRef:string;
    Id:string;

    
    }