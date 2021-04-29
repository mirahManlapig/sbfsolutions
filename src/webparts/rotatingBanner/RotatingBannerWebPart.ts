import { Version } from '@microsoft/sp-core-library';
import { 
  BaseClientSideWebPart, 
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider

} from '@microsoft/sp-webpart-base';
import {IRotatingBannerWebPartProps} from './IRotatingBannerWebPartProps';
import {IRotatingBannerItem,IRotatingBannerList} from './ISPList';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader'; 
import * as strings from 'RotatingBannerWebPartStrings';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';

//import * as $ from 'jquery';
//require('popper.js');
require('jquery');
require('bootstrap');

export default class RotatingBannerWebPart extends BaseClientSideWebPart<IRotatingBannerWebPartProps> {
  protected onInit(): Promise<void> {
    //Add external CSS file from CDN
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/bootstrap.min.css`);
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/fabric.components.min.css`);
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/fabric.min.css`);
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/style.css`);
    SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl + `/SiteAssets/css/CustomStyles.css`);
    //SPComponentLoader.loadScript(this.context.pageContext.web.absoluteUrl + `/SiteAssets/js/jquery-1.9.1.js`);
    //SPComponentLoader.loadScript(this.context.pageContext.web.absoluteUrl + `/SiteAssets/js/popper.min.js`);
    //SPComponentLoader.loadScript(this.context.pageContext.web.absoluteUrl + `/SiteAssets/js/bootstrap.min.js`);
    return super.onInit();
  }

  public render(): void {   
    this.domElement.innerHTML = `
 <section class="carouselbanner">
                    <div id="carouselbanner" class="carousel slide " data-ride="carousel" data-interval="${parseInt(this.properties.sliderTime+"000")}">  
                        <ol class="carousel-indicators" id="sliderDots">
                            
                        </ol>
                        <div class="carousel-inner" id="RotatingBannerID">
                        </div>
                        <a class="carousel-control-prev" href="#carouselbanner" role="button" data-slide="prev"> <i class="ms-Icon ms-Icon--ChevronLeftMed" "="" aria-hidden=" true"></i> <span class="sr-only">Previous</span> </a> <a class="carousel-control-next" href="#carouselbanner" role="button" data-slide="next"> <i class="ms-Icon ms-Icon--ChevronRight" "="" aria-hidden=" true"></i> <span class="sr-only">Next</span> </a>
                        </div>
                        </section>
      `;
      this._renderListDataAsyncRotatingBanner();
  }


  private _getListItemsRotatingBanner():Promise<IRotatingBannerList>{
    var today = new Date();
    var dateFormat = today.toISOString();
    dateFormat = dateFormat.split('T')[0];
    dateFormat = dateFormat + "T00:00:00";
    var query = this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('"+this.properties.ListName+"')/items?$select=Title,BannerDescription,PreviewImageurl,FileLeafRef,Id&$filter=Effectivedate le datetime'"+dateFormat+"' and Expirydate ge datetime'"+dateFormat+"'";
    console.log(query);
    return this.context.spHttpClient.get(query,SPHttpClient.configurations.v1)
    .then((responseListRotatingBanner:SPHttpClientResponse)=>{
      return responseListRotatingBanner.json();
    })
  }

  private _renderListRotatingBanner(Items:IRotatingBannerItem[]):void{
    let RotatingBannerHtml:string='';
    let sliderDotsHtml:string='';
    var slideCount=0;
  if (Items !=null && Items.length>0) {
    Items.forEach((item:IRotatingBannerItem) => {
     
      var className="carousel-item";
      if (slideCount==0) {
        className="carousel-item active";
        sliderDotsHtml+=' <li data-target="#carouselbanner" data-slide-to="'+slideCount+'" class="active"></li>';
      }
      else{
        sliderDotsHtml+=' <li data-target="#carouselbanner" data-slide-to="'+slideCount+'" ></li>';
      }
      
      var BannerDescription=item.BannerDescription;
      if (BannerDescription==null) {
        BannerDescription="";
      }
     
      else{
        if (BannerDescription.length>100) {
          BannerDescription=BannerDescription.substring(0,100)+"...";
        }
      }
      var bannerTitle=item.Title;
      if (bannerTitle==null) {
        bannerTitle="";
      }
      else{
      if(bannerTitle.length>50){
        bannerTitle=bannerTitle.substring(0,50)+"...";
        bannerTitle = '<h5>'+bannerTitle+'</h5>';
      }
    }
      var redirectUrl=item.PreviewImageurl;
      if (redirectUrl==null) {
        //redirectUrl=this.context.pageContext.web.absoluteUrl+"/"+this.properties.ListName+"/Forms/DispForm.aspx?ID="+item.Id;
        redirectUrl="#";
      }
      else {
        redirectUrl=item.PreviewImageurl['Url'];
      }

      
      RotatingBannerHtml+=   
      `<div class="${className}" onclick="window.location.href='${redirectUrl}'"> `+
      '<img src="'+this.context.pageContext.web.absoluteUrl.toString()+"/"+this.properties.ListName+"/"+item.FileLeafRef.toString()+'?RenditionID=5" class="d-block w-100" alt="...">'+
      '<div class="carousel-caption d-none d-sm-block">'+
          '<div class="caption-cnt">'+
              //'<h5>'+bannerTitle+'</h5>'+
              bannerTitle+
              '<p class="d-none d-md-block">'+BannerDescription+'</p>'+
          '</div>'+
      '</div>'+
  '</div>';
  slideCount++;

         
  });
  this.domElement.querySelector('#sliderDots').innerHTML=sliderDotsHtml;
  let quickListContainer:Element=this.domElement.querySelector("#RotatingBannerID");
  quickListContainer.innerHTML=RotatingBannerHtml;
}
else{

  this.domElement.querySelector('#carouselbanner').innerHTML="<h6> No Banner items to display</h6>";
  
}
  
 

  }
private _renderListDataAsyncRotatingBanner(){
  this._getListItemsRotatingBanner().then((Response)=>{
    this._renderListRotatingBanner(Response.value);
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
                  value: "RotatingBanner"
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
