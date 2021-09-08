import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneCustomFieldProps,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  IPropertyPaneDropdownOption,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import {
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';


import * as strings from 'NewsRotatorWebPartStrings';
import NewsRotator from './components/NewsRotator';
import { INewsRotatorProps } from './components/INewsRotatorProps';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {PropertyPaneMessageBar} from './PropertyPaneAutocomplete';
import ListViewService from './Data/ListViewService'; 
import { Spinner } from 'office-ui-fabric-react';

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

  protected onInit(): Promise<void> {  
    return super.onInit().then(() => {  
      ListViewService.setup(this.context); 
    });
  }
  
  public render(): void {
    ReactDom.render(React.createElement('div',
      {style:{minHeight:410, display:'flex', alignItems:'center', justifyContent:'center', backgroundColor:'white'}},
      React.createElement(Spinner,
        {label:'Loading...',styles:{circle:{width:70,height:70}}},)), 
      this.domElement);
    ListViewService.getRotatorNews(this.properties.fixedItems).then(rotatorNews => {
      ListViewService.getAllNews().then(allNews => {
        this.fullList = allNews;
        allNews.forEach((news,i) => {
          this.fullListResult.push({
            key: news.Id,
            text: news.Title
          });
        });
        const element: React.ReactElement<INewsRotatorProps > = React.createElement(
          NewsRotator,
          {
            title: this.properties.title,
            autoplay: this.properties.autoplay,
            fade: this.properties.fade,
            listItems: rotatorNews,
            animationSpeed: this.properties.animationSpeed,
            autoplaySpeed: this.properties.autoplaySpeed,
            fourtElement: this.properties.fourtElement,
            fifthElement: this.properties.fifthElement,
            fullList: this.fullList,
            fixedItems: this.properties.fixedItems
          }
        );
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
      });
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  //@ts-ignore
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private getListItemByKey = (key) => {
    return this.fullListResult.filter(page => page.key === key)[0];
  }

  private setCustomFieldValue = (field,value) => {
    this.properties[field] = value ? value.key : null;
    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    let templateProperty4: any;
    let templateProperty5: any;
    if (this.properties.fixedItems) {
      templateProperty4 = PropertyPaneMessageBar('fourtElement',{
        label: strings.fourtElementFieldLabel,
        onPropertyChange: (event: any, newValue: any) => this.setCustomFieldValue('fourtElement',newValue),
        properties: this.properties,
        value: this.getListItemByKey(this.properties.fourtElement),
        options: this.fullListResult,
        key: 'fourtElement'
      });
      templateProperty5 = PropertyPaneMessageBar('fifthElement',{
        label: strings.fourtElementFieldLabel,
        onPropertyChange: (event: any, newValue: any) => this.setCustomFieldValue('fifthElement',newValue),
        properties: this.properties,
        value: this.getListItemByKey(this.properties.fifthElement),
        options: this.fullListResult,
        key: 'fifthElement'
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
