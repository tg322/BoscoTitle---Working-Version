export interface IIconPropertyPanePropsHost {
    key: string;
    value: any;
    label: string;
    iconColor: any;
    iconBackgroundColor: any;
    onChanged?: (value: any) => void; // eslint-disable-line @typescript-eslint/no-explicit-any
  }
  
  export interface IIconPropertyPanePropsHostState {
    value: any;
    iconColor: any;
    iconBackgroundColor: any;
    isOpen: boolean;
  }