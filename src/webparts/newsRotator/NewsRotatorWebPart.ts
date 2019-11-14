import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'NewsRotatorWebPartStrings';
import NewsRotator from './components/NewsRotator';
import { INewsRotatorProps } from './components/INewsRotatorProps';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

export interface INewsRotatorWebPartProps {
  title: string;
  listItems:any;
  autoplay: boolean;
  fade: boolean;
  animationSpeed: number;
  autoplaySpeed: number;
  fourtElement: any;
  fifthElement: any;
  fullList: any;
  fixedItems: boolean;
}

export default class NewsRotatorWebPart extends BaseClientSideWebPart<INewsRotatorWebPartProps> {

  private listResult;
  private fullList = [];
  private fullListResult = [];
  private listInit = false;
  private fullListInit = false;

  public onInit<T>(): Promise<T> {
    let today = new Date().toISOString();
    let queryRotator = '';
    queryRotator += '$select=ID,Title,BannerImageUrl,FileRef&';
    queryRotator += '$filter=(NewsRotator eq 1) and (PromotedState eq 2) and (FinalApproved eq 1) and (FSObjType eq 0) and (ExpireDate gt \''+today+'\' )&';
    if(this.properties.fixedItems){
      queryRotator += '$top=3&';
    } else {
      queryRotator += '$top=5&';
    }
    queryRotator += '$orderby=FirstPublishedDate desc';
    this._getListData(queryRotator).then((response) => {
      this.listResult = response.value;
      this.listInit = true;
      this.render();
    });
    let queryNews = '';
    queryNews += '$select=ID,Title,BannerImageUrl,FileRef&';
    queryNews += '$orderby=Title asc';
    this._getListData(queryNews).then((response) => {
      this.fullList = response.value;
      response.value.forEach((element,i) => {
        this.fullListResult.push({
          key: element.Id,
          text: element.Title
        });
      });
      this.fullListInit = true;
      this.render();
    });
    return Promise.resolve();
  }
  
  public render(): void {
    
    const element: React.ReactElement<INewsRotatorProps > = React.createElement(
      NewsRotator,
      {
        title: this.properties.title,
        autoplay: this.properties.autoplay,
        fade: this.properties.fade,
        listItems: this.listResult,
        animationSpeed: this.properties.animationSpeed,
        autoplaySpeed: this.properties.autoplaySpeed,
        fourtElement: this.properties.fourtElement,
        fifthElement: this.properties.fifthElement,
        fullList: this.fullList,
        fixedItems: this.properties.fixedItems
      }
    );
    if(this.listInit && this.fullListInit){
      if(this.properties.fixedItems){
        this.fullList.forEach((item,i) => {
          if(item.Id == this.properties.fourtElement) {
            element.props.listItems[3] = item;
          }
          if(item.Id == this.properties.fifthElement) {
            element.props.listItems[4] = item;
          }
        });
      } else {
        element.props.listItems = this.listResult;
      }
      ReactDom.render(element, this.domElement);
    }
  }

  private _getListData(query:string): Promise<any> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/Lists/GetByTitle('Site Pages')/Items?` + query, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    let templateProperty4: any;
    let templateProperty5: any;
    if (this.properties.fixedItems) {
      templateProperty4 = PropertyPaneDropdown('fourtElement', {
        label: strings.fourtElementFieldLabel,
        options: this.fullListResult
      });
      templateProperty5 = PropertyPaneDropdown('fifthElement', {
        label: strings.fifthElementFieldLabel,
        options: this.fullListResult
      });
    } else {
      templateProperty4 = PropertyPaneLabel('emptyLabel', {
        text: ""
      });
      templateProperty5 = PropertyPaneLabel('emptyLabel', {
        text: ""
      });
    }
    
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneToggle('autoplay', {
                  label: strings.AutoplayFieldLabel
                }),
                PropertyPaneSlider('autoplaySpeed', {
                  label: strings.AutoplaySpeedFieldLabel,  
                  min:2000,  
                  max:10000,  
                  value:3000,  
                  showValue:true,  
                  step:500
                }),
                PropertyPaneToggle('fade', {
                  label: strings.FadeFieldLabel
                }),
                PropertyPaneSlider('animationSpeed', {
                  label: strings.AnimationSpeedFieldLabel,  
                  min:100,  
                  max:2000,  
                  value:500,  
                  showValue:true,  
                  step:100
                }),
                PropertyPaneToggle('fixedItems', {
                  label: strings.FixedLastItemsFieldLabel
                }),
                templateProperty4,
                templateProperty5
              ]
            }
          ]
        }
      ]
    };
  }
}
