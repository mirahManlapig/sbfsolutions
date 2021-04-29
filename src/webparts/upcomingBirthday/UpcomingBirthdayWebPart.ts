import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider

} from '@microsoft/sp-webpart-base';
require('bootstrap');

import { IUpcomingBirthdayWebPartProps } from './IUpcomingBirthdayWebPartProps';
import { IUpcomingBirthdayItem, IUpcomingBirthdayList } from './ISPList';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './UpcomingBirthdayWebPart.module.scss';
import * as strings from 'UpcomingBirthdayWebPartStrings';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
//import * as $ from 'jquery';
require('jquery');
require('bootstrap');
//require('moment');

export default class UpcomingBirthdayWebPart extends BaseClientSideWebPart<IUpcomingBirthdayWebPartProps> {

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
                        <section class="birthday-carousel">
                            <h3 class="line-heading"><span class="bluey-purple">Upcoming Birthday</span></h3>
                            <div id="birthday-carousels" class="carousel slide" data-ride="carousel" data-interval="${parseInt(this.properties.sliderTime + "000")}">
                                <ol class="carousel-indicators" id="sliderDots">
                                    
                                </ol>
                                <div class="card">
                                <img src="${this.properties.BirthdayImageUrl}" class="card-img-top" alt="...">
                                <div class="carousel-inner" id="UpcomingBirthdayID">
                                </div>
                                </div>
                                </div>
                                </section>
                    
                
      `;
    this._renderListDataAsyncUpcomingBirthday();
  }
  private _getListItemsUpcomingBirthday(): Promise<IUpcomingBirthdayList> {
    var today = new Date();
    today.setDate(today.getDate() - 1);
    var dateFormat = today.toISOString();
    //var dateFormatTo = new Date().toISOString().split('T')[0];
    var dateFormatTo = new Date(today.getFullYear(), today.getMonth() + 1, 0).toISOString().split('T')[0];
    var dateFormatFrom = new Date(today.getFullYear(), today.getMonth(), 1).toISOString().split('T')[0];
    dateFormatFrom = dateFormatFrom + "T16:00:00";
    dateFormatTo = dateFormatTo + "T15:59:59";
    var query = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + this.properties.ListName + "')/items?$select=Title,Description,picture,Created,Currentdate,ID&$filter=Currentdate ge datetime'" + dateFormatFrom + "' and Currentdate le datetime'" + dateFormatTo + "'";
    //console.log('BDay From:' + dateFormatFrom);
    //console.log('BDay To:' + dateFormatTo);
    console.log(query);

    return this.context.spHttpClient.get(query, SPHttpClient.configurations.v1)
      .then((responseListUpcomingBirthday: SPHttpClientResponse) => {
        return responseListUpcomingBirthday.json();
      })
  }

  private _renderListUpcomingBirthday(Items: IUpcomingBirthdayItem[]): void {
    let UpcomingBirthdayHtml: string = '';
    let sliderDotsHtml: string = '';
    var slideCount = 0;
    if (Items != null && Items.length > 0) {

      Items.forEach((item: IUpcomingBirthdayItem) => {
        //console.log(item.Currentdate);
        var className = "carousel-item";
        if (slideCount == 0) {
          className = "carousel-item active";
          sliderDotsHtml += ' <li data-target="#birthday-carousels" data-slide-to="' + slideCount + '" class="active"></li>';
        }
        else {
          sliderDotsHtml += ' <li data-target="#birthday-carousels" data-slide-to="' + slideCount + '" ></li>';
        }

        var profileImageUrl = item.picture;
        if (profileImageUrl == null) {
          profileImageUrl = this.context.pageContext.web.absoluteUrl + "/SiteAssets/img/profile.png";
        }
        else {
          profileImageUrl = item.picture['Url']
        }
        var displayFormUrl = this.context.pageContext.web.absoluteUrl + "/Lists/" + this.properties.ListName + "/DispForm.aspx?ID=" + item.ID + "&Source=" + this.context.pageContext.web.absoluteUrl;
        UpcomingBirthdayHtml +=
          '<div class="' + className + '">' +
          '<div class="birthday-carousels-content">' +
          // '<div class="card">' +
          '<div class="card-body">' +
          '<a href="' + displayFormUrl + '" style="color: black;text-decoration: none;"><div class="media"> <img src=' + profileImageUrl + ' class="align-self-start mr-3" alt="...">' +
          '<div class="media-body">' +
          '<h5 class="mt-0 bday-name" style="margin-top:1em !important;">' + item.Title + '</h5>' +
          // '<h5 class="mt-0 bday-date">' + this.getForamttedDate(item.Currentdate) + '</h5>' +
          '<h5 class="mt-0 bday-date"></h5>' +
          '</div>' +
          '</div></a>' +
          // '<p class="card-text">' + item.Description + '</p>' +
          '<p class="card-text"></p>' +
          // '</div>' +
          '</div>' +
          '</div>' +
          '</div>';


        slideCount++;


      });
      this.domElement.querySelector('#sliderDots').innerHTML = sliderDotsHtml;
      let quickListContainer: Element = this.domElement.querySelector("#UpcomingBirthdayID");
      quickListContainer.innerHTML = UpcomingBirthdayHtml;
    }
    else {
      this.domElement.querySelector('#birthday-carousels').innerHTML = "<h6>No Birthdays today</h6>";
    }



  }
  private _renderListDataAsyncUpcomingBirthday() {
    this._getListItemsUpcomingBirthday().then((Response) => {
      this._renderListUpcomingBirthday(Response.value);
    })

  }

  private getForamttedDate(currentDate) {
    //var BirthDate = currentDate.substring(0, 10)
    var year = currentDate.substring(0, 4);
    var month = currentDate.substring(5, 7);
    var date = currentDate.substring(8, 10);
    //console.log("Year" + year);
    //console.log("Month" + month);
    //console.log("Day" + date);
    var formattedDate = new Date(year as number - 1, month as number - 1, date as number);

    //console.log("formattedDate : " + formattedDate);
    var BirthDate = this.addDays(formattedDate, 1);
    //console.log("BirthDate : " + BirthDate);
    var arrayMonths = ['Jan', 'Feb', 'Mar,', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    var finalDateString = BirthDate.getDate() + " " + arrayMonths[BirthDate.getMonth()] + " " + BirthDate.getFullYear();
    return finalDateString;

  }

  private addDays(date, days) {
    const copy = new Date(Number(date))
    copy.setDate(date.getDate() + days)
    return copy
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
                  value: "UpcomingBirthday"
                }),
                PropertyPaneTextField('BirthdayImageUrl', {
                  label: "Default Birthday Image",
                  value: this.context.pageContext.web.absoluteUrl + `/SiteAssets/img/bday_party.png`
                }),
                PropertyPaneSlider('sliderTime', {
                  label: 'Slider Time in Seconds',
                  min: 1,
                  max: 10,
                  value: 5
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
