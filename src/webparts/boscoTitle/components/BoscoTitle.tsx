import * as React from 'react';
import styles from './BoscoTitle.module.scss';
import { IBoscoTitleProps } from './IBoscoTitleProps';
// import * as ReactIcons from '@fluentui/react-icons-mdl2';
import QuickLinkContainer from './ReactFragments/QuickLinkContainer'

//import fluent ui icons and figure out how to generate the selected Icon!!

function BoscoTitle(props: IBoscoTitleProps){
  const {
    image1,
    image1Position,
    pageTitle,
    pageTitleColor,
    pageParagraph,
    quickLink1Icon,
    quickLink1IconColor,
    quickLink1IconContainerColor,
    quickLink1Title,
    quickLink1Url,
    quickLink1NewTab,
    quickLink2Icon,
    quickLink2IconColor,
    quickLink2IconContainerColor,
    quickLink2Title,
    quickLink2Url,
    quickLink2NewTab,
    quickLink3Icon,
    quickLink3IconColor,
    quickLink3IconContainerColor,
    quickLink3Title,
    quickLink3Url,
    quickLink3NewTab,
    quickLink4Icon,
    quickLink4IconColor,
    quickLink4IconContainerColor,
    quickLink4Title,
    quickLink4Url,
    quickLink4NewTab
  } = props;
  // const IconComponent = (ReactIcons as any)[quickLink1Icon];

  let quickLinks: { [keys: string]: any; } = {};

  if(quickLink1Title){
    quickLinks.quickLink1 = {
      Icon: quickLink1Icon,
      Title: quickLink1Title,
      Url: quickLink1Url,
      NewTab: quickLink1NewTab,
      IconColor: quickLink1IconColor,
      IconBackgroundColor: quickLink1IconContainerColor
    }
  }
  if(quickLink2Title){
    quickLinks.quickLink2 = {
      Icon: quickLink2Icon,
      Title: quickLink2Title,
      Url: quickLink2Url,
      NewTab: quickLink2NewTab,
      IconColor: quickLink2IconColor,
      IconBackgroundColor: quickLink2IconContainerColor
    }
  }
  if(quickLink3Title){
    quickLinks.quickLink3 = {
      Icon: quickLink3Icon,
      Title: quickLink3Title,
      Url: quickLink3Url,
      NewTab: quickLink3NewTab,
      IconColor: quickLink3IconColor,
      IconBackgroundColor: quickLink3IconContainerColor
    }
  }
  if(quickLink4Title){
    quickLinks.quickLink4 = {
      Icon: quickLink4Icon,
      Title: quickLink4Title,
      Url: quickLink4Url,
      NewTab: quickLink4NewTab,
      IconColor: quickLink4IconColor,
      IconBackgroundColor: quickLink4IconContainerColor
    }
  }

  // console.log(quickLinks);

  return(
    <section className={`${styles.titleBody}`}>
      <div className={`${styles.backgroundImageContainer}`} id={image1 && image1.label ? image1.label : ''} style={{backgroundImage: `url(${image1 && image1.blob ? image1.blob : ''})`, backgroundSize:'cover', backgroundPosition:image1Position}}>
        <div className={`${styles.backgroundOverlay}`}></div>
          <div className={`${styles.titleContainer}`}>
            <h1 style={{color: pageTitleColor}}>{pageTitle}</h1>
            <h2 style={{color: pageTitleColor}}>{pageParagraph}</h2>
          </div>
          <div className={`${styles.quickLinksContainer}`}>
            {Object.keys(quickLinks).map((key) => {
              const quickLink = quickLinks[key];
              return (

                <QuickLinkContainer key={quickLink.Title} Title={quickLink.Title} Url={quickLink.Url} NewTab={quickLink.NewTab} Icon={quickLink.Icon} IconColor={quickLink.IconColor} IconBackgroundColor={quickLink.IconBackgroundColor}/>
                
              );
            })}
          </div>
        
      </div>
    </section>
  );
}

export default BoscoTitle
