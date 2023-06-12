import {
  IPropertyPaneCustomFieldProps,
} from '@microsoft/sp-property-pane';

export interface IBgUploadPropertyPaneProps {
  key: string;
  context?: any;
  fileName: string;
  acceptsType?: string;
  overwrite?: boolean;
  libraryName: string;
  value: any;
  label: string;
  onChanged?: (value: any) => void; // eslint-disable-line @typescript-eslint/no-explicit-any
}

export interface IBgUploadPropertyPanePropsInternal extends IBgUploadPropertyPaneProps, IPropertyPaneCustomFieldProps { }