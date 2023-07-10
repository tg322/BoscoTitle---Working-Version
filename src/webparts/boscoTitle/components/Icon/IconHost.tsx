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
import { EditIcon } from '@fluentui/react-icons-mdl2';
import { PanelType } from '@fluentui/react/lib/Panel';
import styles from './Icon.module.scss';
import * as ReactIcons from '@fluentui/react-icons-mdl2';




export default class PropertyFieldIconHost extends React.Component<
  IIconPropertyPanePropsHost,
  IIconPropertyPanePropsHostState
> {

  constructor(props: IIconPropertyPanePropsHost) {
    super(props);
    this.state = {
      value: this.props.value,
      isOpen: false,
      iconColor: this.props.iconColor,
      iconBackgroundColor: this.props.iconBackgroundColor
    };

    this.selectedIcon = this.selectedIcon.bind(this);
    
  }
  componentDidUpdate(prevProps: Readonly<IIconPropertyPanePropsHost>, prevState: Readonly<IIconPropertyPanePropsHostState>, snapshot?: any): void {
      if(prevProps.value != this.props.value){
        this.setState({value: this.props.value}, () => {
          this.props.onChanged(this.state.value);
      }); 
      }
  }

  openPanel = () => {
    this.setState({ isOpen: true });
  }
  
  dismissPanel = () => {
    this.setState({ isOpen: false });
  }

  selectedIcon(iconName:any) {
    this.setState({value: iconName}, () => {
      this.props.onChanged(this.state.value);
      this.setState({ isOpen: false });
  }); 
  }


  public render(): React.ReactElement<IIconPropertyPanePropsHost> {
    const IconComponent = (ReactIcons as any)[this.state.value];
    
    return (
    <>
      <div className={`${styles.selectedIconContainer}`}>
        <div className={`${styles.selectedIconSquare}`} style={{backgroundColor: this.props.iconBackgroundColor }}>
          <IconComponent className={`${styles.selectedIcon}`} style={{color: this.props.iconColor }}/>
        </div>
        
      </div>
      <div className={`${styles.selectIconContainer}`} onClick={this.openPanel}>
        
        <div className={`${styles.pencilIcon}`}>
          <EditIcon/>
        </div>
        <div className={`${styles.selectIconText}`}>
          <p>Select Icon</p>
        </div>
      </div>
      
      <div>
      <Panel
        isOpen={this.state.isOpen}
        onDismiss={this.dismissPanel}
        type={PanelType.extraLarge}
        closeButtonAriaLabel="Close"
        headerText="Select Icon"
      >
        <IconSelectScreen onClick={this.selectedIcon} />
        
      </Panel>
    </div>
      
      
    </>
    );
  }
}


