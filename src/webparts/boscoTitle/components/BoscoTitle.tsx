import * as React from 'react';
import styles from './BoscoTitle.module.scss';
import { IBoscoTitleProps } from './IBoscoTitleProps';


function BoscoTitle(props: IBoscoTitleProps){
  const {
    image1,
    image1Position
  } = props;
  return(
    <section className={`${styles.titleBody}`}>
        
        <div className={`${styles.backgroundImageContainer}`} id={image1 && image1.label ? image1.label : ''} style={{backgroundImage: `url(${image1 && image1.blob ? image1.blob : ''})`, backgroundSize:'cover', backgroundPosition:image1Position}}>

          <div className={`${styles.quickLinksContainer}`}>
            <div><p>Hello this is a test</p></div>
            <div><p>Hello this is a test</p></div>
            <div><p>Hello this is a test</p></div>
            <div><p>Hello this is a test</p></div>
          </div>

        </div>

        {/* <div id={image2 && image2.label ? image2.label : ''} style={{backgroundImage: `url(${image2 && image2.blob ? image2.blob : ''})`, backgroundSize:'cover', backgroundPosition:image2Position}}>
          
        </div> */}
        
      </section>
  );
}

export default BoscoTitle
