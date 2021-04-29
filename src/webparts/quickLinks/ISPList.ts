

/**
 * @interface
 * Defines a SharePoint QuickLinks list items
 */
export interface IQuickLinksList{
    value:IQuickLinksItem[];
}

/**
 * @interface
 * Defines a SharePoint QuickLinks list item
 */
export interface IQuickLinksItem{
    Title:string;
    Imageurl:string;
    Redirecturl:string;
   
}
