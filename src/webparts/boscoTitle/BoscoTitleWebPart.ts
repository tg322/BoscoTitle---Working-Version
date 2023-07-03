import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'BoscoTitleWebPartStrings';
import BoscoTitle from './components/BoscoTitle';
import { IBoscoTitleProps } from './components/IBoscoTitleProps';
import { PropertyFieldBgUpload } from './components/backgroundUpload/BgUploadPropertyPane';
import { spfi, SPFx as spSPFx } from "@pnp/sp";
import '@pnp/sp/folders';
import { folderFromServerRelativePath } from '@pnp/sp/folders';

export interface IBoscoTitleWebPartProps {
  description: string;
  image1: any;
  image1Position: string;
  image2: any;
  image2Position: string;
  context: any;
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


  // private pageTitle: string = '';
  private sp: any;

  public render(): void {
    const element: React.ReactElement<IBoscoTitleProps> = React.createElement(
      BoscoTitle,
      {
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

    this.sp = spfi().using(spSPFx(this.context));

    const folder = await folderFromServerRelativePath(this.sp.web, 'Shared Documents/Bosco Title/'+this.context.pageContext.site.id).select('Exists')();

    if(!folder.Exists){
      await this.sp.web.folders.addUsingPath('Shared Documents/Bosco Title/'+this.context.pageContext.site.id);
      // const files = await this.sp.web.getFolderByServerRelativePath('Shared Documents/Bosco Title/'+this.context.pageContext.site.id).files();
      // console.log(files);
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
  
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldBgUpload("image1", {
                  key: "image1",
                  label: "image1",
                  value: this.properties.image1,
                  context: this.context,
                  fileName: this.fileNames.image1FileName,
                  libraryName: 'Shared Documents/Bosco Title/'+this.context.pageContext.site.id
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
                
                // PropertyFieldBgUpload("image2", {
                //   key: "image2",
                //   label: "image2",
                //   value: this.properties.image2,
                //   context: this.context,
                //   fileName: this.fileNames.image2FileName,
                //   libraryName: 'Shared Documents/Bosco Title'+this.context.pageContext.site.id
                // }),
                // PropertyPaneDropdown('image2Position', {
                //  label: "Image position",
                //  options: [
                //   { key: 'bottom', text: 'Bottom'},
                //   { key: 'center', text: 'Center' },
                //   { key: 'left', text: 'Left' },
                //   { key: 'right', text: 'Right'},
                //   { key: 'top', text: 'Top'}
                // ],
                // selectedKey: this.properties.image2Position
                  
                // })
              ]
            }
          ]
        }
      ]
    };
  }
}
