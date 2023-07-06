export interface IIconPropertyPanePropsHost {
    key: string;
    value: any;
    label: string;
    onChanged?: (value: any) => void; // eslint-disable-line @typescript-eslint/no-explicit-any
  }
  
  export interface IIconPropertyPanePropsHostState {
    value: any;
    isOpen: boolean;
  }