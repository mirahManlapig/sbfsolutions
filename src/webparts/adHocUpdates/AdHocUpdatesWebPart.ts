import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,  
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AdHocUpdatesWebPart.module.scss';
import * as strings from 'AdHocUpdatesWebPartStrings';
import {IAnnouncementsItem,IAnnouncementsList} from './ISPList';
import {IAnnouncementsWebPartProps} from './IAnnouncementsWebPartProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
//import * as $ from 'jquery';
//require('popper.js');
require('jquery');
require('bootstrap');

export default class AdHocUpdatesWebPart extends BaseClientSideWebPart<IAnnouncementsWebPartProps> {

  protected onInit(): Promise<void> {
    //Add external CSS file from CDN
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/bootstrap.min.css`);
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/fabric.components.min.css`);
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/fabric.min.css`);
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/style.css`);
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/CustomStyles.css`);
    
    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `

    <section class="">
    <div id="carouselnews-ad" class="carousel slide" data-interval="${parseInt(this.properties.sliderTime+"000")}">
        <ol class="carousel-indicators" id="sliderAnnouncementsDots">
            
        </ol>
        <div class="carousel-inner" id="AnnouncementsID">
        </div>
        </div>
       <a class="carousel-control-prev" href="#carouselnews-ad" role="button" data-slide="prev"> <i class="ms-Icon ms-Icon--ChevronLeftMed" "="" aria-hidden=" true"></i> <span class="sr-only">Previous</span> </a> <a class="carousel-control-next" href="#carouselnews-ad" role="button" data-slide="next"> <i class="ms-Icon ms-Icon--ChevronRight" "="" aria-hidden=" true"></i> <span class="sr-only">Next</span> </a>
       </div>  
       </section>
     `;
     this._renderListDataAsyncAnnouncements();
  }

  private _getListItemsAnnouncements():Promise<IAnnouncementsList>{

    var today = new Date();
    var dateFormat = today.toISOString();
    dateFormat = dateFormat.split('T')[0];
    dateFormat = dateFormat + "T00:00:00";
    var query = this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('"+this.properties.ListName+"')/items?$select=Title,Description,Effectivedate,Expirydate,Created,ID&$filter=Effectivedate le datetime'"+dateFormat+"' and Expirydate ge datetime'"+dateFormat+"'";
    console.log(query);
    return this.context.spHttpClient.get(query,SPHttpClient.configurations.v1)
    .then((responseListAnnouncements:SPHttpClientResponse)=>{
      return responseListAnnouncements.json();
    })
  }

  private _renderListAnnouncements(Items:IAnnouncementsItem[]):void{
    let AnnouncementsHtml:string='';
    let sliderAnnouncementsDotsHtml:string='';
    var slideCount=0;
    if (Items !=null && Items.length>0) {
    Items.forEach((item:IAnnouncementsItem) => {
     
      var className="carousel-item";
      if (slideCount==0) {
        className="carousel-item active";
        sliderAnnouncementsDotsHtml+=' <li data-target="#carouselnews-ad" data-slide-to="'+slideCount+'" class="active"></li>';
      }
      else{
        sliderAnnouncementsDotsHtml+=' <li data-target="#carouselnews-ad" data-slide-to="'+slideCount+'" ></li>';
      }
      var annDescription=item.Description;
      if (annDescription==null) {
        annDescription="";
      }
      var displayFormUrl=this.context.pageContext.web.absoluteUrl+"/Lists/"+this.properties.ListName+"/DispForm.aspx?ID="+item.ID+"&Source="+this.context.pageContext.web.absoluteUrl;
      AnnouncementsHtml+=   
      
      '<div class="'+className+'">'+
      '<div class="carouselnews-content">'+
          /*'<a href="'+displayFormUrl+'"><div class="news-time"> <i class="ms-Icon ms-Icon--CalendarMirrored" aria-hidden="true"></i> '+this.getForamttedDate(item.Created)+'</div></a>'+*/
          '<a href="'+displayFormUrl+'" style="text-decoration: none;color: black;"><h5>'+item.Title+'</h5></a>'+
          '<p>'+annDescription+'</p>'+
      '</div>'+
  '</div>';

     
  slideCount++;

         
  });
  this.domElement.querySelector('#sliderAnnouncementsDots').innerHTML=sliderAnnouncementsDotsHtml;
  let quickListContainer:Element=this.domElement.querySelector("#AnnouncementsID");
  quickListContainer.innerHTML=AnnouncementsHtml;
}
else{
  
  this.domElement.querySelector("#AnnouncementsID").innerHTML="<h6>No Announcements to display</h6>";
  $('#carouselnews-ad').children('a.carousel-control-prev').hide();
  $('#carouselnews-ad').children('.carousel-control-next').hide();
}
  
 

  }
private _renderListDataAsyncAnnouncements(){
  this._getListItemsAnnouncements().then((Response)=>{
    this._renderListAnnouncements(Response.value);

  })

}



  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('ListName', {
                  label: "ListName",
                  value: "AdHocUpdates"
                }),
                PropertyPaneTextField('webPartTitle', {
                  label: "Webpart Title",
                  value: "Ad-Hoc Updates"
                }),
                PropertyPaneSlider('sliderTime',{
                  label:'Slider Time in Seconds',
                  min:1,
                  max:10,
                  value:5
                 }),
              ]
            }
          ]
        }
      ]
    };
  }
}
