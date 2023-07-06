import * as React from "react";
import {
  IIconPropertyPanePropsHost,
  IIconPropertyPanePropsHostState
} from "./IIconPropertyPanePropsHost";
// import { graphfi, SPFx as graphSPFx} from "@pnp/graph";
import IconSelectScreen from './IconSelectScreen';
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { Panel } from "office-ui-fabric-react";
import { PanelType } from '@fluentui/react/lib/Panel';




export default class PropertyFieldIconHost extends React.Component<
  IIconPropertyPanePropsHost,
  IIconPropertyPanePropsHostState
> {

  constructor(props: IIconPropertyPanePropsHost) {
    super(props);
    this.state = {
      value: this.props.value,
      isOpen: false

      
    };

  }

  openPanel = () => {
    this.setState({ isOpen: true });
  }
  
  dismissPanel = () => {
    this.setState({ isOpen: false });
  }

  // handleIconSelectToggle = () => {
  //   this.setState(prevState => ({ iconSelectVisible: !prevState.iconSelectVisible }));
  // }

  
  
  public render(): React.ReactElement<IIconPropertyPanePropsHost> {
    
    return (
    <>
      <div>
        <button onClick={this.openPanel}></button>
      </div>
      
      
      <div>
      <Panel
        isOpen={this.state.isOpen}
        onDismiss={this.dismissPanel}
        type={PanelType.extraLarge}
        closeButtonAriaLabel="Close"
        headerText="Sample panel"
      >
        <IconSelectScreen />
        
      </Panel>
    </div>
      
      
    </>
    );
  }
}


