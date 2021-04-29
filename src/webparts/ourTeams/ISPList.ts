/**
 * @interface
 * Defines a SharePoint OurTeams list items
 */
export interface IOurTeamsList{
    value:IOurTeamsItem[];
    }

    /**
 * @interface
 * Defines a SharePoint OurTeams list item
 */
    export interface IOurTeamsItem{
    Title:string;
    Description:string;
    Departmentname:string;
    Imageurl:string;
    Created:string;
    Id:string;
    Redirecturl:
      {
         Description:string,
         Url:string
      };
    }