import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class DataHandler {

  async getFormDigestValue(context: any) {
    const response = await context.spHttpClient.post(`${context.pageContext.web.absoluteUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, {});
    const responseJSON = await response.json();
    const formDigestValue = responseJSON.FormDigestValue;
    return formDigestValue;
  }

  async uploadFileToSP(file: File, context: any, libraryName: string, overwrite: boolean, fileName?:string): Promise<any> {
    //Get formDigestValue
    const formDigestValue = await this.getFormDigestValue(context);
    
    let reader = new FileReader();
    return new Promise((resolve, reject) => {
      reader.onload = (event: any) => {
        let blob = new Blob([event.target.result], { type: file.type });
  
        const url = `${context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${libraryName}')/files/add(overwrite=${overwrite}, url='${fileName && fileName ? fileName : file.name}')`;
  
        const headers = {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': file.type,
          'odata-version': '',
          'X-RequestDigest': formDigestValue
        };
  
        context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
          body: blob,
          headers: headers
        })
        .then((response: SPHttpClientResponse) => {
          if(response.ok) {
            response.json().then((fileData: any) => {
              resolve(fileData);
            });
          }
          else {
            reject(new Error(`Error uploading file: ${response.statusText}`));
          }
        });
      };
      reader.readAsArrayBuffer(file);
    });
  }

  async deleteFileFromSP(context: any, libraryName: string, fileName: string): Promise<void> {
    const formDigestValue = await this.getFormDigestValue(context);

    const url = `${context.pageContext.web.absoluteUrl}/_api/web/getfilebyserverrelativeurl('/${libraryName}/${fileName}')`;

    const headers = {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
        'X-RequestDigest': formDigestValue,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
    };

    context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: headers
    })
    .then((response: SPHttpClientResponse) => {
        // Handle response
    });
}

async checkFolderExistsInSP(context: any, libraryName: string, folderName: string): Promise<boolean> {
  const formDigestValue = await this.getFormDigestValue(context);
  return new Promise((resolve, reject) => {

  const url = `${context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('/${libraryName}/${folderName}')/Exists`;

  const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=verbose',
      'odata-version': '',
      'X-RequestDigest': formDigestValue
  };

  context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: headers
  })
  .then((response: SPHttpClientResponse) => {
    if(response.ok) {
      response.json().then((exists: boolean) => {
        resolve(exists);
      });
    }
    else {
      reject(new Error(`Error locating folder: ${response.statusText}`));
    }
  });

});

}



async createFolderInSP(context: any, folderLocation: string, folderName:string): Promise<any>{

  // const formDigestValue = await this.getFormDigestValue(context);
  
  let serverRelativeUrl = `${context.pageContext.web.serverRelativeUrl}${folderLocation}/${folderName}`;

  let requestUrl = `${context.pageContext.web.absoluteUrl}/_api/web/folders`;

  context.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, {
      headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': ''
      },
      body: JSON.stringify({
          '__metadata': { 'type': 'SP.Folder' },
          'ServerRelativeUrl': serverRelativeUrl
      })
  })
  .then((response: SPHttpClientResponse) => {
      if (response.ok) {
          console.log(`Folder '${folderName}' created successfully!`);
      } else {
          console.log(`Failed to create folder. Status: ${response.status} (${response.statusText})`);
      }
  })
  .catch((error: any) => {
      console.error(`Error creating folder: ${error}`);
  });
  }




  async checkFileExistsInSP(context: any, filePath: string, fileName:string): Promise<any> {

    return new Promise((resolve, reject) => {
      console.log(filePath);
    const url = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${filePath}')/Files`;
  
    context.spHttpClient.get(url, SPHttpClient.configurations.v1,)
    .then((response: SPHttpClientResponse) => {
      if(response.ok) {
        response.json().then((files: any) => {
          let matchingFiles = files.value.filter((file: any) =>{
            
            const cleanFileName = file.Name.substring(0, file.Name.indexOf("."));
            
            return cleanFileName === fileName;
          });
          
          if (matchingFiles.length > 0) {
            resolve(true + ' ' + fileName + matchingFiles.Name);
          } else {
            resolve(false + ' ' + fileName)
          }
          
        });
        
      }
      else {
        reject(new Error(`Error locating folder: ${response.statusText}`));
      }
    });
  
  });
  
  }


}