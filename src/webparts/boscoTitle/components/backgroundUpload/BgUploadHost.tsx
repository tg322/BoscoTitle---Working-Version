import * as React from "react";
import {
  IBgUploadPropertyPanePropsHost,
  IBgUploadPropertyPanePropsHostState
} from "./IBgUploadPropertyPanePropsHost";
import { Icon, Spinner } from "office-ui-fabric-react";
import styles from './BgUpload.module.scss';
import Modal from './Modal';
import { spfi, SPFx as spSPFx } from "@pnp/sp";
// import { graphfi, SPFx as graphSPFx} from "@pnp/graph";

import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { IFileAddResult } from "@pnp/sp/files";



export default class PropertyFieldBgUploadHost extends React.Component<
  IBgUploadPropertyPanePropsHost,
  IBgUploadPropertyPanePropsHostState
> {

  private sp: any;

  constructor(props: IBgUploadPropertyPanePropsHost) {
    super(props);

    //if a custom state is needed, such as the isVisible, add to IBgUploadPropertyPanePropsHostState to avoid getting errors in groupFields

    const value = this.props.value;

    this.state = {
      value: value,
      //If value contains an image, set isVisible to true, else, false. Allows persistence over refresh
      isVisible: value != null,
      isUploading: false,
      modalVisible: false
      
    };

    this.sp = spfi().using(spSPFx(this.props.context));
  

  }
  

  // class set to async to allow use of await function to hold code while functions run. Prevents bugs with uploading and removing images not being completed when setting states and showing previews.
  private handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    //If <input> event target contains a file and the length is longer than 0 *when using for the first time 'files' does not exist, however when the image is removed the 'files' array persists but it is empty hence the two checks*
    if (event.target.files && event.target.files.length > 0) {
      //set file as the result of the uploaded image
      const file = event.target.files[0];
      
      //If value already contains an image *uploading an image then clicking upload again instead of clicking remove then uploading* run handleFileRemove (remove file from sharepoint)
      if(this.state.value != null){
        //wait for this function to finish before continuing with the code, sending file name allows this function to locate the old file to then remove it
        // await this.handleFileRemove(this.state.value.fileName);

        // let graph = graphfi().using(graphSPFx(this.context));

        // let sp = spfi().using(spSPFx(this.context));

        await this.sp.web.getFolderByServerRelativePath(this.props.libraryName).files.getByUrl(this.state.value.fileName).delete();


      }
      //Run upload file function, wait for this function to finish
      await this.handleFileUpload(file);
      //Tell sharepoint the value property has changed with that of the value state. (state is the react side, property is the main webpart, set property to state to sync the two)
      this.props.onChanged(this.state.value);

    }
  };

handleFileUpload = async (file: File) => {

  this.setState({ isUploading: true, isVisible: false });

  let result: IFileAddResult;
  const fileType = file.name.slice(file.name.indexOf('.'));
  try{
    result = await this.sp.web.getFolderByServerRelativePath(this.props.libraryName).files.addUsingPath(this.props.fileName+fileType, file, { Overwrite: true });

    let fileObject: { [keys: string]: any; } = {};

    fileObject.fileName = this.props.fileName+fileType;
  
    fileObject.blob = this.props.context.pageContext.web.absoluteUrl + encodeURI(result.data.ServerRelativeUrl) + `?UUID=${new Date().getTime() + Math.floor(Math.random() * 1000000000)}`;
  
    fileObject.label = this.props.label;

    this.setState({value: fileObject});

    this.props.onChanged(this.state.value);
  }
  catch (error){
    console.error('Error uploading file:', error);
  }
  
  this.setState({ isUploading: false, isVisible: true });

  this.props.onChanged(this.state.value);
  
}

handleModalToggle = () => {
  this.setState(prevState => ({ modalVisible: !prevState.modalVisible }));
}

//Handle click is the function added to the remove button, this is run as an async function so that it is completed when called using await which means the code will wait for this function to fully complete before continuing
handleClick = async () => {
    //Run the handleFileRemove function as await, insert the fileName from the value prop
    this.handleModalToggle();
    await this.handleFileRemove(this.props.value.fileName);
    //Once the file has been removed, set the value state to null (remove the image that is to be removed from the webpart) and hide the preview by setting isVisible to false
    this.setState({value: null, isVisible: false});
    //Tell sharepoint theres nothing in the props
    this.props.onChanged(this.state.value);
    //Clear the <input> tag, this will keep the file if not cleared and can cause many many bugs
    if (this.uploadInputRef.current) {
      this.uploadInputRef.current.value = "";
    }
    
};

//The async handleFileRemove function that removes the file from sharepoint
handleFileRemove = async (fileName:any) => {
    //Create an instance of DataHandler to access the data handling functions
    
    //Begin the async remove function
    try {
      //Run await the deleteFileFromSP function
      await this.sp.web.getFolderByServerRelativePath(this.props.libraryName).files.getByUrl(fileName).delete();
    } catch (error) {
      console.error('Error deleting file:', error);
    }
}

  //Create reference to the input field, this can now be used in code elsewhere such as when we clear the contents of the <input> tag in handleClick
  private uploadInputRef = React.createRef<HTMLInputElement>();
  
  public render(): React.ReactElement<IBgUploadPropertyPanePropsHost> {
    
    return (
    <>
      <div className={`${styles.imageUploadContainer}`}>
        { !this.state.isUploading && 
        <div className={`${styles.imageUploadButtonContainer}`}>
          <label>
          <input ref={this.uploadInputRef} type="file" id={`${this.props.label}`} name="message" accept={`${this.props.acceptsType && this.props.acceptsType ? this.props.acceptsType : 'image/*'}`} onChange={this.handleFileChange}></input>
            <div>
              <Icon iconName="FileImage"></Icon>
              <p>Upload</p>
            </div>
          </label>
          { this.state.isVisible && <div  onClick={this.handleModalToggle}>
              <Icon iconName="Delete"></Icon>
            <p>Remove</p>
          </div>}
        </div>}
      { this.state.isVisible && <div className={`${styles.imagePreviewContainer}`}>
        <img className={`${styles.imagePreview}`} src={this.state.value.blob} alt="Preview" />
      </div> }
      

      { this.state.isUploading && <div style={{display: `flex`, flexDirection: `column`, height: `100px`, justifyContent:`center`}}>
        
        <Spinner style={{flexDirection: `column`}} label="Uploading Image to SharePoint..." />
      </div> }
      
      </div>
      { this.state.modalVisible && <Modal titleAction="Delete image from SharePoint" prompt="Are you sure you wish to delete this image?" image={this.state.value.blob} action={this.handleClick} closeModal={this.handleModalToggle}/> }
    </>
    );
  }
}



