import { IPropertyPaneCustomFieldProps, IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

import {ComboBox} from 'office-ui-fabric-react/lib/index';

import TextField from '@material-ui/core/TextField';
import Autocomplete from '@material-ui/lab/Autocomplete';

interface IPropertyFieldMessageBarProps {  
    label: string;  
    onPropertyChange(event: any, newValue: any): void;  
    properties: any;
    value: any;
    options: any[];  
    key?: string;  
}

interface IPropertyFieldMessageBarPropsInternal extends IPropertyPaneCustomFieldProps {  
    label: string;   
    targetProperty: string; 
    options: any[]; 
    onRender(elem: HTMLElement): void; // Method to render data (in this case reactjs is used)  
    onDispose(elem: HTMLElement): void; // Method to delete the object  
    onPropertyChange(event: any, newValue: any): void;  
    properties: any; // set of properties (holds values of property collection defined using IPropertyFieldMessageBarProps) 
    value: any; 
    key: string;  
}

export class PropertyFieldMessageBarBuilder implements IPropertyPaneField<IPropertyFieldMessageBarPropsInternal> {  
  
    //Properties defined by IPropertyPaneField  
    public properties: IPropertyFieldMessageBarPropsInternal;  
    public targetProperty: string;  
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;  
    
    //Custom properties  
    private label: string;  
    private onPropertyChange: (event: any, newValue: any) => void;  
    private customProperties: any;
    private options: any[];  
    private key: string;
    private value: any;  

    public dispose = () => {

    }
      
    // Initializes the varialbes or methods  
    public constructor(usertargetProperty: string, userproperties: IPropertyFieldMessageBarPropsInternal) {  
      this.render = this.render.bind(this);  
      this.targetProperty = userproperties.targetProperty;  
      this.properties = userproperties;  
      this.label = userproperties.label;
      this.options = userproperties.options;  
      this.properties.onDispose = this.dispose;  
      this.properties.onRender = this.render;  
      this.onPropertyChange = userproperties.onPropertyChange;  
      this.customProperties = userproperties.properties;  
      this.key = userproperties.key;
      this.value = userproperties.value;  
    }  
    
    // Renders the data  
    private render(elem: HTMLElement): void {  
        /*const element = React.createElement(ComboBox,{
            label: this.label,
            key: this.key,
            allowFreeform: false,
            autoComplete: "on",
            options: this.options
        });*/

        const getInput = (params) => {
            return React.createElement(TextField,{...params, label: this.label, variant:'outlined'});
        };
        const element = React.createElement(Autocomplete,{
            id:this.key,
            options: this.options,
            value: this.value,
            onChange: this.onPropertyChange,
            getOptionLabel: (option) => option['text'],
            style: { width: 300, marginTop: 20 },
            renderInput: getInput
        });
        ReactDOM.render(element,elem);
    }  
    
}

export function PropertyPaneMessageBar(targetProperty: string, properties: IPropertyFieldMessageBarProps): IPropertyPaneField<IPropertyFieldMessageBarPropsInternal> {  
  
    // Builds the property based on the custom data  
    var newProperties: IPropertyFieldMessageBarPropsInternal = {  
      label: properties.label,  
      targetProperty: targetProperty,  
      onPropertyChange: properties.onPropertyChange,  
      properties: properties.properties,
      value: properties.value,  
      onDispose: null,  
      onRender: null, 
      options: properties.options, 
      key: properties.key  
    };  
      
    // Initialize and render the properties  
    return new PropertyFieldMessageBarBuilder(targetProperty, newProperties);  
}