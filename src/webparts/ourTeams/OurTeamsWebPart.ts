import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField

} from '@microsoft/sp-webpart-base';
import { IOurTeamsWebPartProps } from './IOurTeamsWebPartProps';
import { IOurTeamsItem, IOurTeamsList } from './ISPList';

import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'OurTeamsWebPartStrings';

import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
//import * as $ from 'jquery';
//import * as caroufredsel from 'caroufredsel';
//import 'jqueryOld';

require('jqueryOld');
//require('popper.js');
//require('jquery');
//require('bootstrap');
require('carouFredSel');

export default class OurTeamsWebPart extends BaseClientSideWebPart<IOurTeamsWebPartProps> {

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
   
            <div class="">
                <h3 class="line-heading"><span class="greeny-blue">Our Teams</span></h3>
                <div class="">
                    <section class="ourteam-carousel">
                       <div class="carouselourteam-content">

                       <div id="ccarousel" class="OurTeamsID">
                       </div>
                       <div id="pager"></div>
                       </div>
                       </section>
                       </div>
                       </div>
                      
                  
    
    `;
    this._renderListDataAsyncOurTeams();
  }

  private _getListItemsOurTeams(): Promise<IOurTeamsList> {
    var query = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.properties.ListName + "')/items?$select=Id,Title,Description,Departmentname,Imageurl,Created,Redirecturl&$orderby=Departmentname asc";
    console.log(query);
    return this.context.spHttpClient.get(query, SPHttpClient.configurations.v1)
      .then((responseListOurTeams: SPHttpClientResponse) => {
        return responseListOurTeams.json();
      })
  }

  private _renderListOurTeams(Items: IOurTeamsItem[]): void {
    let OurTeamsHtml: string = '';
    if (Items != null && Items.length > 0) {


      Items.forEach((item: IOurTeamsItem) => {

        var OurteamsDescription = item.Description;
        if (OurteamsDescription == null) {
          OurteamsDescription = "";
        }

        else {
          OurteamsDescription = item.Description;
        }

        var OurteamsTitle = item.Title;
        if (OurteamsTitle == null) {
          OurteamsTitle = "";
        }
        else {
          OurteamsTitle = item.Title;
        }

        var OurteamsImageurl = item.Imageurl;
        if (OurteamsImageurl == null) {
          OurteamsImageurl = this.context.pageContext.web.absoluteUrl + "/SiteAssets/img/team.jpg";
        }
        else {
          OurteamsImageurl = item.Imageurl['Url'];
        }

        var Redirecturl = item.Redirecturl.Url;
        if (Redirecturl == null || Redirecturl.trim() == '') {
          Redirecturl = "#";
        }
        else {
          Redirecturl = item.Redirecturl.Url;
        }


        OurTeamsHtml += '<div class="card team-card" style="height: 250px !important;" data-arg1="">' +
          /* '<ul class="list-group list-group-flush">'+
               '<li class="list-group-item"><a	href="'+item.Redirecturl['Url']+'">'+item.Departmentname+'</a></li>'+
          '</ul>'+ */
          '<a	href="' + Redirecturl + '">'+
          '<img src="' + OurteamsImageurl + '" class="card-img-top" alt="...">' +
          '</a>' +
          //'<div class="card-body">' +
          //'<h5 class="card-title"><a	href="' + this.context.pageContext.web.absoluteUrl + "/Lists/" + this.properties.ListName + "/DispForm.aspx?ID=" + item.Id + "&Source=" + this.context.pageContext.web.absoluteUrl + '">' + OurteamsTitle + '</a></h5>' +
          //'<h5 class="card-title">' + OurteamsTitle + '</h5>' +
          //'<p class="card-text">' + OurteamsDescription + '</p>' +
          //'</div>' +
          //'<div class="teamtime card-footer"> <i class="ms-Icon ms-Icon--CalendarMirrored" aria-hidden="true"></i> ' + this.getForamttedDate(item.Created) + '</div>' +
          '</div>';


      });

      let quickListContainer: Element = this.domElement.querySelector(".OurTeamsID");
      quickListContainer.innerHTML = OurTeamsHtml;

      this.callOurTeamsCarousel();
    }
    else {
      this.domElement.querySelector(".ourteam-carousel").innerHTML = "<h6>No Departments to display</h6>";
    }


  }
  private _renderListDataAsyncOurTeams() {
    this._getListItemsOurTeams().then((Response) => {
      this._renderListOurTeams(Response.value);
    })

  }

  private getForamttedDate(currentDate) {
    var formattedDate = new Date(currentDate);
    var arrayMonths = ['Jan', 'Feb', 'Mar,', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    var finalDateString = formattedDate.getDate() + " " + arrayMonths[formattedDate.getMonth()] + " " + formattedDate.getFullYear();
    return finalDateString;

  }

  private callOurTeamsCarousel() {
    var _direction = 'left';
    (<any>$('#ccarousel')).carouFredSel({
      direction: _direction,
      responsive: true,
      circular: false,
      items: {
        width: 350,
        height: '100%',
        visible: {
          min: 2,
          max: 5
        }
      },
      pagination: '#pager',
      scroll: {
        items: 1,
        duration: 2000,
        timeoutDuration: 500,
        pauseOnHover: 'immediate',
        onEnd: function (data) {
          _direction = (_direction == 'left') ? 'right' : 'left';
          $(this).trigger('configuration', ['direction', _direction]);
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('ListName', {
                  label: "ListName",
                  value: "OurTeams"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
