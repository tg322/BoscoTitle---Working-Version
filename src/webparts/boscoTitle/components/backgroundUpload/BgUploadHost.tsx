import * as React from "react";
import {
  IBgUploadPropertyPanePropsHost,
  IBgUploadPropertyPanePropsHostState
} from "./IBgUploadPropertyPanePropsHost";
import { Icon, Spinner } from "office-ui-fabric-react";
import styles from '../BoscoTitle.module.scss';
import {DataHandler} from '../DataHandler';
import Modal from '../modal/Modal';





export default class PropertyFieldBgUploadHost extends React.Component<
  IBgUploadPropertyPanePropsHost,
  IBgUploadPropertyPanePropsHostState
> {
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
        await this.handleFileRemove(this.state.value.fileName);
      }
      //Run upload file function, wait for this function to finish
      await this.handleFileUpload(file);
      //Tell sharepoint the value property has changed with that of the value state. (state is the react side, property is the main webpart, set property to state to sync the two)
      this.props.onChanged(this.state.value);

    }
  };

//Function to run when uploading an image
handleFileUpload = async (file: File) => {
    //Show the uploading to sharepoint spinner and hide the preview and upload/remove buttons
    this.setState({ isUploading: true, isVisible: false });
    //Create a new instance of the DataHandler class from DataHandler.ts, this file contains the upload, remove logic
    let dataHandler = new DataHandler();
    //Using the original filename e.g 'BoscoHeroBackgroundImage.png' grab only the '.png' extention
    const fileType = file.name.slice(file.name.indexOf('.'));
    //begin the async upload function
    try {
      //Apply the call to dataHandler uploadFileToSP function to a variable in order to fetch the files details once its uploaded to the document library
      const fileData = await dataHandler.uploadFileToSP(file, this.props.context, this.props.libraryName, this.props.overwrite && this.props.overwrite ? this.props.overwrite : true, this.props.fileName && this.props.fileName ? this.props.fileName+fileType : file.name);
      //Create an object, we do this because creating a large object for props is complex and can become convuluted easily, so having one 'value' prop and storing an object within it is a simpler approach
      let fileObject: { [keys: string]: any; } = {};
      //Store the fileName from the returned fileData in the fileName key
      fileObject.fileName = fileData.Name;
      //Store the blob from the returned fileData in the blob key. The code this.props.context.pageContext.web.absoluteUrl grabs the main url of the sharepoint site e.g https://stpaulscc.sharepoint.com, we then get the relative URL from the returned fileData
      //which will contain spaces and other characters that can cause errors so we encode this url to replace these special characters such as spaces from ' ' to '%20' which creates a valid url. We also add a timestamp to the end of the URL to
      //solve the caching issues of uploading an image which has been renamed to that of the property fileName, which in our case is 'image1' .png, and can display the previous image stored in the cache since if another .png file is uploaded and renamed
      //to 'image1' the client side will simply pull the cached image, adding the timestamp denotes this as a unique url meaning it will pull the new image each time
      fileObject.blob = this.props.context.pageContext.web.absoluteUrl + encodeURI(fileData.ServerRelativeUrl) + `?UUID=${new Date().getTime() + Math.floor(Math.random() * 1000000000)}`;
      fileObject.label = this.props.label;
      //We set the state of value to the fileObject key value pair object
      this.setState({value: fileObject});
      //Let sharepoint framework know the prop value is now the state value to be used in the main webpart
      this.props.onChanged(this.state.value);

    } catch (error) {
      console.error('Error uploading file:', error);
    }
    //Revert the states back to their defaults, hide the uploading to sharepoint spinner and show the preview and upload/remove buttons
    this.setState({ isUploading: false, isVisible: true });
    //Set the onChanged property again, dont want to remove this because its working and not sure if removing will reveal a bug :(
    // this.props.onChanged(this.state.value);
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
    let dataHandler = new DataHandler();
    //Begin the async remove function
    try {
      //Run await the deleteFileFromSP function
      await dataHandler.deleteFileFromSP(this.props.context, this.props.libraryName, fileName);
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



