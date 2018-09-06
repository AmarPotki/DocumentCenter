import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IWebPartContext
} from '@microsoft/sp-webpart-base';
import PropertyFieldTermSetPickerHost from './PropertyFieldTermSetPickerHost';
import { IPropertyFieldTermSetPickerHostProps } from './IPropertyFieldTermSetPickerHost';
import { IPropertyFieldTermSetPickerPropsInternal, IPropertyFieldTermSetPickerProps, ICheckedTermSets } from './IPropertyFieldTermSetPicker';

/**
 * Represents a PropertyFieldTermSetPicker object
 */
class PropertyFieldTermSetPickerBuilder implements IPropertyPaneField<IPropertyFieldTermSetPickerPropsInternal> {

  // Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldTermSetPickerPropsInternal;

  // Custom properties label: string;
  private label: string;
  private context: IWebPartContext;
  private allowMultipleSelections: boolean = false;
  private initialValues: ICheckedTermSets = [];
  private excludeSystemGroup: boolean = false;
  private panelTitle: string;

  public onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void { }
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;
  private onGetErrorMessage: (value: ICheckedTermSets) => string | Promise<string>;
  private deferredValidationTime: number = 200;

  /**
   * Constructor method
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldTermSetPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _targetProperty;
    this.properties = _properties;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.label = _properties.label;
    this.context = _properties.context;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    this.panelTitle = _properties.panelTitle;

    if (_properties.disabled === true) {
      this.disabled = _properties.disabled;
    }
    if (_properties.deferredValidationTime) {
      this.deferredValidationTime = _properties.deferredValidationTime;
    }
    if (typeof _properties.allowMultipleSelections !== 'undefined') {
      this.allowMultipleSelections = _properties.allowMultipleSelections;
    }
    if (typeof _properties.initialValues !== 'undefined') {
      this.initialValues = _properties.initialValues;
    }
    if (typeof _properties.excludeSystemGroup !== 'undefined') {
      this.excludeSystemGroup = _properties.excludeSystemGroup;
    }
  }

  /**
   * Renders the SPListPicker field content
   */
  private render(elem: HTMLElement, ctx?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
    // Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldTermSetPickerHostProps> = React.createElement(PropertyFieldTermSetPickerHost, {
      label: this.label,
      targetProperty: this.targetProperty,
      panelTitle: this.panelTitle,
      allowMultipleSelections: this.allowMultipleSelections,
      initialValues: this.initialValues,
      excludeSystemGroup: this.excludeSystemGroup,
      context: this.context,
      onDispose: this.dispose,
      onRender: this.render,
      onChange: changeCallback,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.key,
      disabled: this.disabled,
      onGetErrorMessage: this.onGetErrorMessage,
      deferredValidationTime: this.deferredValidationTime
    });

    // Calls the REACT content generator
    ReactDom.render(element, elem);
  }

  /**
   * Disposes the current object
   */
  private dispose(elem: HTMLElement): void {

  }

}

/**
 * Helper method to create a SPList Picker on the PropertyPane.
 * @param targetProperty - Target property the SharePoint list picker is associated to.
 * @param properties - Strongly typed SPList Picker properties.
 */
export function PropertyFieldTermSetPicker(targetProperty: string, properties: IPropertyFieldTermSetPickerProps): IPropertyPaneField<IPropertyFieldTermSetPickerPropsInternal> {
  // Create an internal properties object from the given properties
  const newProperties: IPropertyFieldTermSetPickerPropsInternal = {
    label: properties.label,
    targetProperty: targetProperty,
    panelTitle: properties.panelTitle,
    allowMultipleSelections: properties.allowMultipleSelections,
    initialValues: properties.initialValues,
    excludeSystemGroup: properties.excludeSystemGroup,
    context: properties.context,
    onPropertyChange: properties.onPropertyChange,
    properties: properties.properties,
    onDispose: null,
    onRender: null,
    key: properties.key,
    disabled: properties.disabled,
    onGetErrorMessage: properties.onGetErrorMessage,
    deferredValidationTime: properties.deferredValidationTime
  };
  // Calls the PropertyFieldTermSetPicker builder object
  // This object will simulate a PropertyFieldCustom to manage his rendering process
  return new PropertyFieldTermSetPickerBuilder(targetProperty, newProperties);
}
