import { Version } from '@microsoft/sp-core-library';
import { 
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,

} from '@microsoft/sp-webpart-base';
import {IQuickLinksWebPartProps} from './IQuickLinksWebPartProps';
import {IQuickLinksItem,IQuickLinksList} from './ISPList';
import * as strings from 'QuickLinksWebPartStrings';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './QuickLinksWebPart.module.scss';
//import * as $ from 'jquery';
require('jquery');
require('bootstrap');
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ISPQuickLinksListService, SPQuicklinksListService} from './SPQuickLinksListService';

import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import 'bootstrap';


export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
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
    <section class="qik-link-carousel">
    <h1 class="line-heading "><span class="blue">${this.properties.webPartTitle}</span></h1>
    <div id="qik-link-carousel" class="carousel slide" >
    <ol class="carousel-indicators">
           <!--- <li data-target="#qik-link-carousel" data-slide-to="0" class="active"></li>
            <li data-target="#qik-link-carousel" data-slide-to="1"></li>--->
        </ol>
    <div class="carousel-inner row w-100 mx-auto" role="listbox"  id="QuickLinksID">
    
    </div> 
        <a class="carousel-control-prev" href="#qik-link-carousel" role="button" data-slide="prev"> <i class="ms-Icon ms-Icon--ChevronLeftMed"" aria-hidden=" true"></i> <span class="sr-only">Previous</span> </a>
        
         <a class="carousel-control-next text-faded" href="#qik-link-carousel" role="button" data-slide="next"> <i class="ms-Icon ms-Icon--ChevronRight"" aria-hidden=" true"></i> <span class="sr-only">Next</span> </a>
         
        </div> 
        
        </section> 
      
                      
      `;
      this._renderListDataAsyncQuickLinks();
  }

  private _getListItemsQuickLinks():Promise<IQuickLinksList>{
    var query = this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('"+this.properties.ListName+"')/items?$select=Title,Imageurl,Redirecturl";
    console.log(query);
    return this.context.spHttpClient.get(query,SPHttpClient.configurations.v1)
    .then((responseListQuickLinks:SPHttpClientResponse)=>{
      return responseListQuickLinks.json();
    })
  }

  private _renderListQuickLinks(Items:IQuickLinksItem[]):void{
    let QuickLinksHtml:string='';
    var slideCount=0;
  if (Items !=null && Items.length>0) {
    Items.forEach((item:IQuickLinksItem) => {

      var className="carousel-item col-sm-2";
      if (slideCount==0) {
        className="carousel-item col-sm-2 active";
        //sliderDotsHtml+=' <li data-target="#carouselbanner" data-slide-to="'+slideCount+'" class="active"></li>';
      }
      else{
        //sliderDotsHtml+=' <li data-target="#carouselbanner" data-slide-to="'+slideCount+'" ></li>';
      }


      var QuickImageUrl=item.Imageurl;
      if (QuickImageUrl==null) {
        QuickImageUrl=this.context.pageContext.web.absoluteUrl+"/SiteAssets/img/Image.png";
      }
else{
  QuickImageUrl=item.Imageurl;
}

      QuickLinksHtml+=     
      '<div class="'+className+'">'+
      '<div class="panel panel-default">'+
          '<div class="panel-thumbnail"> <a href="'+item.Redirecturl+'" title="image 1" class="thumb"> <img class="img-fluid mx-auto d-block" src='+QuickImageUrl+' alt="slide 1">'+
                  '<div class="carousel-caption d-block">'+
                      '<p>'+item.Title+'</p>'+
                  '</div>'+
              '</a></div>'+
      '</div>'+
  '</div>';
  slideCount++;

               
  });
  let quickListContainer:Element=this.domElement.querySelector("#QuickLinksID");
  quickListContainer.innerHTML=QuickLinksHtml;
  this.callQuickLinksslider();
}
else{
  this.domElement.querySelector("#qik-link-carousel").innerHTML="<h6>No Quick Links items to display<h6>";

}

  }
private _renderListDataAsyncQuickLinks(){
  this._getListItemsQuickLinks().then((Response)=>{
    this._renderListQuickLinks(Response.value);
  })

}
private callQuickLinksslider(){
var itemsInSlideProp=this.properties.itemsInSlide;
  $('#qik-link-carousel').on('slide.bs.carousel', function(e) { 
    var g = $(e.relatedTarget);
     var idx = g.index(); 
     var itemsPerSlide =6;
     var totalItems = $('.qik-link-carousel .carousel-item').length; 
     if (idx >= totalItems - (itemsPerSlide - 1)) {
        var it = itemsPerSlide - (totalItems - idx);
         for (var i = 0; i < it; i++){
            if (e.direction == "left")
             {
                $('.qik-link-carousel .carousel-item').eq(i).appendTo('.qik-link-carousel .carousel-inner');
             }
            else 
            {
              $('.qik-link-carousel .carousel-item').eq(0).appendTo('.qik-link-carousel .carousel-inner'); 
            }
           
          }
         }
        });
}

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:""
          },
          groups: [
            {
              groupName: "",
              groupFields: [
                PropertyPaneTextField('ListName', {
                  label: "ListName",
                  value: "QuickLinks"
                }),
                PropertyPaneTextField('webPartTitle', {
                  label: "WebPart Title",
                  value: "Quick Links"
                }),
                PropertyPaneSlider('itemsInSlide', {
                  label: "Number of items in slide",
                  min:5,
                  max:10,
                  value:6
                })
               
              ]
            }
          ]
        }
      ]
    };
  }
}
