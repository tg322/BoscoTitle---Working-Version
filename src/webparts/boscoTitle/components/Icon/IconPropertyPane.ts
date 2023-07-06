import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
} from '@microsoft/sp-property-pane';

import { IIconPropertyPaneProps, IIconPropertyPanePropsInternal } from './IIconPropertyPaneProps';
import PropertyFieldIconHost from './IconHost';

//Constructs the property pane, including all the properties and the react that is to be rendered.

class PropertyFieldIconBuilder implements IPropertyPaneField<IIconPropertyPanePropsInternal> {
  public targetProperty: string;
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public properties: IIconPropertyPanePropsInternal;


  private _onChangeCallback: (targetProperty?: string, newValue?: any) => void; // eslint-disable-line @typescript-eslint/no-explicit-any
  
  //Property constructor

  public constructor(_targetProperty: string, _properties: IIconPropertyPanePropsInternal) {
    this.targetProperty = _targetProperty;
    this.properties = _properties;
    this.properties.onChanged = this._onChanged.bind(this);
    this.properties.onRender = this._render.bind(this);
    this.properties.onDispose = this._dispose.bind(this);
  }

  private _render(elem: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void { // eslint-disable-line @typescript-eslint/no-explicit-any

    const props: IIconPropertyPaneProps = <IIconPropertyPaneProps>this.properties;

    const element = React.createElement(PropertyFieldIconHost, {
      ...props
    });

    ReactDOM.render(element, elem);

    if (changeCallback) {
      this._onChangeCallback = changeCallback;
    }
  }

  private _dispose(elem: HTMLElement): void {
    ReactDOM.unmountComponentAtNode(elem);
  }

  private _onChanged(value: any): void {
    if (this._onChangeCallback) {
      this._onChangeCallback(this.targetProperty, value);
    }

  }

}

//Function that is used within the group fields to create this custom field, this takes the arguments constructed above and uses return new PropertyFieldIconBuilder to create an entirely new instance of the file upload property pane.
//this way the same custom property pane can be reused simply by calling the function in the group fields and assigning a separate variable for the value, key value and label

export function PropertyFieldIcon(targetProperty: string, properties: IIconPropertyPaneProps): IPropertyPaneField<IIconPropertyPanePropsInternal> {
  
  return new PropertyFieldIconBuilder(targetProperty, {
    ...properties,
    onChanged: properties.onChanged,
    onRender: null,
    onDispose: null
  });
}