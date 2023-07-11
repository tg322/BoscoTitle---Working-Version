import * as React from 'react';
import * as ReactIcons from '@fluentui/react-icons-mdl2';
// import { ITestProps } from './ITestProps';
import styles from '../BoscoTitle.module.scss';

function QuickLinkContainer(props: any){
  
  const IconComponent = (ReactIcons as any)[props.Icon];
  return(
    // main backround for modal
    <a className={`${styles.quickLinkMainContainer}`} key={props.Title} href={props.Url} target={`${props.NewTab && props.NewTab ? '_blank' : '_self'}`}>
        {/* modal box */}
        <div className={`${styles.quickLinkIconContainer}`} style={{backgroundColor: props.IconBackgroundColor}}>
          <IconComponent className={`${styles.quickLinkIcon}`} style={{color: props.IconColor, alignSelf: 'center'}}/>
        </div>
        <div className={`${styles.quickLinkTitleContainer}`}>
            <p>{props.Title}</p>
        </div>
    </a>
  );
}

export default QuickLinkContainer