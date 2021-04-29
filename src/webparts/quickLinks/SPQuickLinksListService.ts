import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {IQuickLinksItem,IQuickLinksList} from './ISPList';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import {IQuickLinksWebPartProps} from './IQuickLinksWebPartProps';
/**
 * @interface
 * Service interface definition
 */
export interface ISPQuickLinksListService {
    /**
     * @function
     * Gets the QuickLinks from a SharePoint list
     */
    getIQuickLinksItems(listId: string): Promise<IQuickLinksList>;
  }

export class SPQuicklinksListService implements ISPQuickLinksListService{

    private context: IWebPartContext;
    private props: IQuickLinksWebPartProps;
  
    /**
     * @function
     * Service constructor
     */
    constructor(_props: IQuickLinksWebPartProps, pageContext: IWebPartContext){
        this.props = _props;
        this.context = pageContext;
    }


     /**
   * @function
   * Gets the QuickLinks from a SharePoint list
   */

   public getIQuickLinksItems(queryUrl: string):Promise<IQuickLinksList>{
    return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        return response.json().then((responseFormated: any) => {
            var formatedResponse: IQuickLinksList = { value: []};
            //Fetchs the Json response to construct the final items list
            responseFormated.value.map((object: any, i: number) => {
                var spListItem: IQuickLinksItem = {
                  
                  'Title':object["Title"],
                  'Imageurl': object["Imageurl"],
                  'Redirecturl':object["Redirecturl"],
                  
                };
                formatedResponse.value.push(spListItem);
            });
            return formatedResponse;
        });
    }) as Promise<IQuickLinksList>;
   }
}