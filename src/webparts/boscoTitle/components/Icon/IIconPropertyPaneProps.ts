import {
  IPropertyPaneCustomFieldProps,
} from '@microsoft/sp-property-pane';

export interface IIconPropertyPaneProps {
  key: string;
  value: any;
  label: string;
  iconColor: any;
  iconBackgroundColor: any
  onChanged?: (value: any) => void; // eslint-disable-line @typescript-eslint/no-explicit-any
}

export interface IIconPropertyPanePropsInternal extends IIconPropertyPaneProps, IPropertyPaneCustomFieldProps { }