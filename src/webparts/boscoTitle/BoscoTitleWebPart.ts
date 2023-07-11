import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, PropertyPaneButton, PropertyPaneButtonType, PropertyPaneDropdown, PropertyPaneTextField, PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'BoscoTitleWebPartStrings';
import BoscoTitle from './components/BoscoTitle';
import PnPTelemetry from "@pnp/telemetry-js";
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker'
import { IBoscoTitleProps } from './components/IBoscoTitleProps';
import { PropertyFieldBgUpload } from './components/backgroundUpload/BgUploadPropertyPane';
import { PropertyFieldIcon } from './components/Icon/IconPropertyPane';
import { SPFx as spSPFx } from "@pnp/sp";
import '@pnp/sp/folders';
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import { Web } from '@pnp/sp/webs';
import { folderFromServerRelativePath } from '@pnp/sp/folders';


export interface IBoscoTitleWebPartProps {
  description: string;
  image1: any;
  image1Position: string;
  image2: any;
  image2Position: string;
  context: any;
  quickLink1Icon: any;
  quickLink1IconColor: any;
  quickLink1IconContainerColor: any;
  quickLink1Title: string;
  quickLink1Url: string;
  quickLink1NewTab: boolean;
  quickLink2Icon: any;
  quickLink2IconColor: any;
  quickLink2IconContainerColor: any;
  quickLink2Title: string;
  quickLink2Url: string;
  quickLink2NewTab: boolean;
  quickLink3Icon: any;
  quickLink3IconColor: any;
  quickLink3IconContainerColor: any;
  quickLink3Title: string;
  quickLink3Url: string;
  quickLink3NewTab: boolean;
  quickLink4Icon: any;
  quickLink4IconColor: any;
  quickLink4IconContainerColor: any;
  quickLink4Title: string;
  quickLink4Url: string;
  quickLink4NewTab: boolean;
  pageTitle: string;
  pageTitleColor: any;
  pageParagraph: string;
}

interface IFileNames {
  [key: string]: string;
}

export default class BoscoTitleWebPart extends BaseClientSideWebPart<IBoscoTitleWebPartProps> {

  

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private fileNames: IFileNames = {
      image1FileName: 'image1',
      image2FileName: 'image2'
  };

  private rootweb: any;

  public render(): void {
    const element: React.ReactElement<IBoscoTitleProps> = React.createElement(
      BoscoTitle,
      {
        pageParagraph: this.properties.pageParagraph,
        pageTitleColor: this.properties.pageTitleColor,
        pageTitle: this.properties.pageTitle,
        quickLink4NewTab: this.properties.quickLink4NewTab,
        quickLink4Url: this.properties.quickLink4Url,
        quickLink4Title: this.properties.quickLink4Title,
        quickLink4IconContainerColor: this.properties.quickLink4IconContainerColor,
        quickLink4IconColor: this.properties.quickLink4IconColor,
        quickLink4Icon:this.properties.quickLink4Icon,
        quickLink3NewTab: this.properties.quickLink3NewTab,
        quickLink3Url: this.properties.quickLink3Url,
        quickLink3Title: this.properties.quickLink3Title,
        quickLink3IconContainerColor: this.properties.quickLink3IconContainerColor,
        quickLink3IconColor: this.properties.quickLink3IconColor,
        quickLink3Icon:this.properties.quickLink3Icon,
        quickLink2NewTab: this.properties.quickLink2NewTab,
        quickLink2Url: this.properties.quickLink2Url,
        quickLink2Title: this.properties.quickLink2Title,
        quickLink2IconContainerColor: this.properties.quickLink2IconContainerColor,
        quickLink2IconColor: this.properties.quickLink2IconColor,
        quickLink2Icon:this.properties.quickLink2Icon,
        quickLink1NewTab: this.properties.quickLink1NewTab,
        quickLink1Url: this.properties.quickLink1Url,
        quickLink1Title: this.properties.quickLink1Title,
        quickLink1IconContainerColor: this.properties.quickLink1IconContainerColor,
        quickLink1IconColor: this.properties.quickLink1IconColor,
        quickLink1Icon:this.properties.quickLink1Icon,
        context: this.context,
        image2Position: this.properties.image2Position,
        image2: this.properties.image2,
        image1Position: this.properties.image1Position,
        image1: this.properties.image1,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {

    const telemetry = PnPTelemetry.getInstance();
    telemetry.optOut();

    if(!this.properties.quickLink1Icon){
      this.properties.quickLink1Icon = 'GlobeIcon'
    }
    if(!this.properties.quickLink2Icon){
      this.properties.quickLink2Icon = 'GlobeIcon'
    }
    if(!this.properties.quickLink3Icon){
      this.properties.quickLink3Icon = 'GlobeIcon'
    }
    if(!this.properties.quickLink4Icon){
      this.properties.quickLink4Icon = 'GlobeIcon'
    }

    try{

      this.rootweb = Web(window.location.origin).using(spSPFx(this.context));

    }catch(error){
      console.log('Error accessing rootWeb: ' + error);
    }
    
    try{
      const mainFolder = await folderFromServerRelativePath(this.rootweb, 'Shared Documents/'+this.context.manifest.alias).select('Exists')();
      if(!mainFolder.Exists){
        try{
          await this.rootweb.folders.addUsingPath('Shared Documents/'+this.context.manifest.alias);
        }catch(error){
          console.log('onInit creating main folder error: ' + error);
        }
       
      }
    }catch(error){
      console.log('onInit checking existance of main folder error: ' + error);
    }
    
    try{
      const siteFolder = await folderFromServerRelativePath(this.rootweb, 'Shared Documents/'+this.context.manifest.alias+'/'+this.context.pageContext.site.id).select('Exists')();

      if(!siteFolder.Exists){
        try{
          await this.rootweb.folders.addUsingPath('Shared Documents/'+this.context.manifest.alias+'/'+this.context.pageContext.site.id);
        }catch(error){
          console.log('onInit creating site folder error: ' + error);
        }
      }
    }catch(error){
      console.log('onInit checking existance of site folder error: ' + error);
    }
    


    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  private resetQuickLinks(quickLink:string){
    if(quickLink === 'QuickLink1'){
      this.properties.quickLink1Icon = 'GlobeIcon';
      this.properties.quickLink1NewTab = null;
      this.properties.quickLink1Title = '';
      this.properties.quickLink1Url = '';
      this.properties.quickLink1IconColor = '#323130';
      this.properties.quickLink1IconContainerColor= '#d3d3d3';
    }else if(quickLink === 'QuickLink2'){
      this.properties.quickLink2Icon = 'GlobeIcon';
      this.properties.quickLink2NewTab = null;
      this.properties.quickLink2Title = '';
      this.properties.quickLink2Url = '';
      this.properties.quickLink2IconColor = '#323130';
      this.properties.quickLink2IconContainerColor= '#d3d3d3';
    }else if(quickLink === 'QuickLink3'){
      this.properties.quickLink3Icon = 'GlobeIcon';
      this.properties.quickLink3NewTab = null;
      this.properties.quickLink3Title = '';
      this.properties.quickLink3Url = '';
      this.properties.quickLink3IconColor = '#323130';
      this.properties.quickLink3IconContainerColor= '#d3d3d3';
    }else if(quickLink === 'QuickLink4'){
      this.properties.quickLink4Icon = 'GlobeIcon';
      this.properties.quickLink4NewTab = null;
      this.properties.quickLink4Title = '';
      this.properties.quickLink4Url = '';
      this.properties.quickLink4IconColor = '#323130';
      this.properties.quickLink4IconContainerColor= '#d3d3d3';
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

 

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const backgroundImage = {
      
      groupName: 'Background Image',
      isCollapsed: true,
              groupFields: [
      PropertyFieldBgUpload("image1", {
        key: "image1",
        label: "image1",
        value: this.properties.image1,
        context: this.context,
        fileName: this.fileNames.image1FileName,
        libraryName: 'Shared Documents/'+this.context.manifest.alias+'/'+this.context.pageContext.site.id
      }),
      PropertyPaneDropdown('image1Position', {
        label: "Image position",
        options: [
         { key: 'center', text: 'Center'},
         { key: 'left', text: 'Left' },
         { key: 'right', text: 'Right' },
         { key: 'top', text: 'Top'},
         { key: 'bottom', text: 'Bottom'}
       ],
       selectedKey: this.properties.image1Position
       })
    ]
    }
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: 'Add Page Background, Title and Quick Links'
          },
          groups: [
            backgroundImage,
            {
              groupName: 'Page Title',
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('pageTitle', {
                  label: 'Enter Page Title',
                  value: this.properties.pageTitle
                }),
                PropertyPaneTextField('pageParagraph', {
                  label: 'Add a Page Paragraph',
                  multiline: true,
                  value: this.properties.pageParagraph
                }),
                PropertyFieldColorPicker('pageTitleColor', {
                  label: 'Title & Paragraph Colour',
                  selectedColor: this.properties.pageTitleColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'pageTitleColor'
                })
              ]
            }
          ]
        },
        {
          displayGroupsAsAccordion: true,
          header: {
            description: 'Add Quick Links'
          },
          groups: [
            {
              groupName: 'Quick Link 1',
              isCollapsed: true,
              groupFields: [
                
                PropertyPaneButton('quickLink1Reset', {
                  text: 'Reset Quick Link',
                  icon: 'Delete',
                  buttonType: PropertyPaneButtonType.Command,
                  onClick: () => this.resetQuickLinks('QuickLink1')
                }),
                PropertyFieldIcon("quickLink1Icon", {
                  key: "quickLink1Icon",
                  label: "quickLink1Icon",
                  value: this.properties.quickLink1Icon,
                  iconColor: this.properties.quickLink1IconColor,
                  iconBackgroundColor: this.properties.quickLink1IconContainerColor
                }),
                PropertyFieldColorPicker('quickLink1IconColor', {
                  label: 'Icon Colour',
                  selectedColor: this.properties.quickLink1IconColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'quickLink1IconColor'
                }),
                PropertyFieldColorPicker('quickLink1IconContainerColor', {
                  label: 'Icon Background Colour',
                  selectedColor: this.properties.quickLink1IconContainerColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'quickLink1IconContainerColor'
                }),
                PropertyPaneTextField('quickLink1Title', {
                  label: 'Enter Title',
                  value: this.properties.quickLink1Title
                }),
                PropertyPaneTextField('quickLink1Url', {
                  label: 'Enter Url',
                  value: this.properties.quickLink1Url
                }),
                PropertyPaneToggle('quickLink1NewTab', {
                  label: 'Open in new tab?',
                  checked: this.properties.quickLink1NewTab
                })
              ]
            }, 
            {
              groupName: 'Quick Link 2',
              isCollapsed: true,
              groupFields: [
                PropertyPaneButton('quickLink1Reset', {
                  text: 'Reset Quick Link',
                  icon: 'Delete',
                  buttonType: PropertyPaneButtonType.Command,
                  onClick: () => this.resetQuickLinks('QuickLink2')
                }),
                PropertyFieldIcon("quickLink2Icon", {
                  key: "quickLink2Icon",
                  label: "quickLink2Icon",
                  value: this.properties.quickLink2Icon,
                  iconColor: this.properties.quickLink2IconColor,
                  iconBackgroundColor: this.properties.quickLink2IconContainerColor
                }),
                PropertyFieldColorPicker('quickLink2IconColor', {
                  label: 'Icon Colour',
                  selectedColor: this.properties.quickLink2IconColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'quickLink2IconColor'
                }),
                PropertyFieldColorPicker('quickLink2IconContainerColor', {
                  label: 'Icon Background Colour',
                  selectedColor: this.properties.quickLink2IconContainerColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'quickLink2IconContainerColor'
                }),
                PropertyPaneTextField('quickLink2Title', {
                  label: 'Enter Title',
                  value: this.properties.quickLink2Title
                }),
                PropertyPaneTextField('quickLink2Url', {
                  label: 'Enter Url',
                  value: this.properties.quickLink2Url
                }),
                PropertyPaneToggle('quickLink2NewTab', {
                  label: 'Open in new tab?',
                  checked: this.properties.quickLink2NewTab
                })
              ]
            },
            {
              groupName: 'Quick Link 3',
              isCollapsed: true,
              groupFields: [
                PropertyPaneButton('quickLink1Reset', {
                  text: 'Reset Quick Link',
                  icon: 'Delete',
                  buttonType: PropertyPaneButtonType.Command,
                  onClick: () => this.resetQuickLinks('QuickLink3')
                }),
                PropertyFieldIcon("quickLink3Icon", {
                  key: "quickLink3Icon",
                  label: "quickLink3Icon",
                  value: this.properties.quickLink3Icon,
                  iconColor: this.properties.quickLink3IconColor,
                  iconBackgroundColor: this.properties.quickLink3IconContainerColor
                }),
                PropertyFieldColorPicker('quickLink3IconColor', {
                  label: 'Icon Colour',
                  selectedColor: this.properties.quickLink3IconColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'quickLink3IconColor'
                }),
                PropertyFieldColorPicker('quickLink3IconContainerColor', {
                  label: 'Icon Background Colour',
                  selectedColor: this.properties.quickLink3IconContainerColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'quickLink3IconContainerColor'
                }),
                PropertyPaneTextField('quickLink3Title', {
                  label: 'Enter Title',
                  value: this.properties.quickLink3Title
                }),
                PropertyPaneTextField('quickLink3Url', {
                  label: 'Enter Url',
                  value: this.properties.quickLink3Url
                }),
                PropertyPaneToggle('quickLink3NewTab', {
                  label: 'Open in new tab?',
                  checked: this.properties.quickLink3NewTab
                })
              ]
            },
            {
              groupName: 'Quick Link 4',
              isCollapsed: true,
              groupFields: [
                PropertyPaneButton('quickLink1Reset', {
                  text: 'Reset Quick Link',
                  icon: 'Delete',
                  buttonType: PropertyPaneButtonType.Command,
                  onClick: () => this.resetQuickLinks('QuickLink4')
                }),
                PropertyFieldIcon("quickLink4Icon", {
                  key: "quickLink4Icon",
                  label: "quickLink4Icon",
                  value: this.properties.quickLink4Icon,
                  iconColor: this.properties.quickLink4IconColor,
                  iconBackgroundColor: this.properties.quickLink4IconContainerColor
                }),
                PropertyFieldColorPicker('quickLink4IconColor', {
                  label: 'Icon Colour',
                  selectedColor: this.properties.quickLink4IconColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'quickLink4IconColor'
                }),
                PropertyFieldColorPicker('quickLink4IconContainerColor', {
                  label: 'Icon Background Colour',
                  selectedColor: this.properties.quickLink4IconContainerColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'quickLink4IconContainerColor'
                }),
                PropertyPaneTextField('quickLink4Title', {
                  label: 'Enter Title',
                  value: this.properties.quickLink4Title
                }),
                PropertyPaneTextField('quickLink4Url', {
                  label: 'Enter Url',
                  value: this.properties.quickLink4Url
                }),
                PropertyPaneToggle('quickLink4NewTab', {
                  label: 'Open in new tab?',
                  checked: this.properties.quickLink4NewTab
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
