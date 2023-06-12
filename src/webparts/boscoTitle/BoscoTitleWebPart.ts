import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {DataHandler} from './components/DataHandler';
import * as strings from 'BoscoTitleWebPartStrings';
import BoscoTitle from './components/BoscoTitle';
import { IBoscoTitleProps } from './components/IBoscoTitleProps';
import { PropertyFieldBgUpload } from './components/backgroundUpload/BgUploadPropertyPane';

export interface IBoscoTitleWebPartProps {
  description: string;
  image1: any;
  image1Position: string;
  image2: any;
  image2Position: string;
  context: any;
}

export default class BoscoTitleWebPart extends BaseClientSideWebPart<IBoscoTitleWebPartProps> {

  

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private fileNames = {
    image1FileName: 'image1',
    image2FileName: 'image2'
    
  };

  private pageTitle: string = '';

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

  protected onInit(): Promise<void> {
    this.pageTitle = this.context.pageContext.site.serverRequestPath.slice(this.context.pageContext.site.serverRequestPath.lastIndexOf('/')+1).substring(0, this.context.pageContext.site.serverRequestPath.slice(this.context.pageContext.site.serverRequestPath.lastIndexOf('/')+1).indexOf("."));
    

    console.log(this.pageTitle);
  
    this.checkFolderExists(this.pageTitle).then(async result => {
      if(result !== null && typeof result === 'object' && 'value' in result){
        if(result.value === false){
          await this.createFolder(this.pageTitle);

          let promises = Object.keys(this.fileNames).map((fileName: string) => {
            let filePath = `Shared Documents/Bosco Title/${this.pageTitle}`;
            return this.checkFile(filePath, fileName);
          });

          Promise.all(promises).then(results => {
            // 'results' is an array of results corresponding to each file check.
            results.forEach((result, index) => {
                console.log(`Check result for file ${Object.keys(this.fileNames)[index]}: ${result}`);
            });
        }).catch(error => {
            console.error("An error occurred while checking the files:", error);
        });

          // console.log(promises);

          await this.checkFile('Shared Documents/Bosco Title/'+this.pageTitle, this.fileNames.image1FileName).then(async result => {
            console.log(result)
          });
        }else if(result.value === true){
          console.log('Folder Exists!');
          await this.checkFile('Shared Documents/Bosco Title/'+this.pageTitle, this.fileNames.image1FileName).then(async result => {
            console.log(result)
          });
        }
      }
      
    });

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  checkFolderExists = async (folderName:string) => {
    //Create an instance of DataHandler to access the data handling functions
    let dataHandler = new DataHandler();
    //Begin the async remove function
    let spResponse:any = '';
    try {
      //Run await the deleteFileFromSP function
      
      console.log(folderName);
      spResponse = await dataHandler.checkFolderExistsInSP(this.context, 'Shared Documents/Bosco Title', folderName);
    } catch (error) {
      console.error('Error deleting file:', error);
    }
    return spResponse;
}

createFolder = async (folderName:string) => {
  //Create an instance of DataHandler to access the data handling functions
  let dataHandler = new DataHandler();
  //Begin the async remove function
  let spResponse:any = '';
  try {
    //Run await the deleteFileFromSP function
    
    spResponse = await dataHandler.createFolderInSP(this.context, 'Shared Documents/Bosco Title', folderName);
  } catch (error) {
    console.error('Error creating file:', error);
  }
  return spResponse;
}

checkFile = async (filePath:string, fileName:string) => {

  let dataHandler = new DataHandler();

  try {
    //Run await the deleteFileFromSP function
    
    await dataHandler.checkFileExistsInSP(this.context, filePath, fileName);
  } catch (error) {
    console.error('Error creating file:', error);
  }
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
                  libraryName: 'Shared Documents/Bosco Title/'+this.pageTitle
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
                  
                }),
                PropertyFieldBgUpload("image2", {
                  key: "image2",
                  label: "image2",
                  value: this.properties.image2,
                  context: this.context,
                  fileName: this.fileNames.image2FileName,
                  libraryName: 'Shared Documents/Bosco Title'
                }),
                PropertyPaneDropdown('image2Position', {
                 label: "Image position",
                 options: [
                  { key: 'center', text: 'Center'},
                  { key: 'left', text: 'Left' },
                  { key: 'right', text: 'Right' },
                  { key: 'top', text: 'Top'},
                  { key: 'bottom', text: 'Bottom'}
                ],
                selectedKey: this.properties.image2Position
                  
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
