import * as React from 'react';
import * as ReactIcons from '@fluentui/react-icons-mdl2';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import styles from './Icon.module.scss';
initializeIcons();

function IconSelectScreen (props:any){

  const [isShown, setIsShown] = React.useState(null);

  function handleSelect(iconName:any) {
    props.onClick(iconName);            
  }

  const classes = mergeStyleSets({
    cell: {
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      margin: '30px',
      float: 'left',
      height: '80px',
      width: '80px',
      cursor: 'pointer',
    },
    icon: {
      fontSize: '30px',
    },
    code: {
      background: '#f2f2f2',
      borderRadius: '4px',
      padding: '4px',
      fontSize: '14px',
    },
    navigationText: {
      width: 100,
      margin: '0 5px',
    },
  });

  const icons = Object.keys(ReactIcons).reduce((acc: React.FC[], exportName) => {
    if ((ReactIcons as any)[exportName]?.displayName) {
      acc.push((ReactIcons as any)[exportName] as React.FunctionComponent);
    }
    return acc;
  }, []);

  return(
    
  <div className={`${styles.modalContainer}`}>
    <div className={`${styles.iconContainer}`}>
      <div>
        {icons
          .map((Icon: React.FunctionComponent<ReactIcons.ISvgIconProps>) => (
            <div key={Icon.displayName} className={classes.cell} onMouseEnter={() => setIsShown(Icon.displayName)}
            onMouseLeave={() => setIsShown(null)} onClick={() => handleSelect(Icon.displayName)}>
              {/*
                Provide an `aria-label` for screen reader users if the icon is not accompanied by
                text that conveys the same meaning.
              */}
              <Icon aria-label={Icon.displayName?.replace('Icon', '')} className={classes.icon}  />
              <br />
              {isShown === Icon.displayName && (
              <code className={classes.code}>{Icon.displayName}</code>
              )}
            </div>
          ))}
      </div>
    </div>
  </div>

  );
}

export default IconSelectScreen