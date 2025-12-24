import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './PnPContext.module.scss';
import type { IPnPContextProps } from './IPnPContextProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { usePnPjs } from './PnPjsContext';

const PnPContext: React.FC<IPnPContextProps> = (props) => {
  const {
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName
  } = props;

  const { sp } = usePnPjs();
  const [webTitle, setWebTitle] = useState<string>("Loading...");

  useEffect(() => {
    const fetchWebTitle = async () => {
      try {
        const web = await sp.web();
        setWebTitle(web.Title);
      } catch (err) {
        console.error("Error fetching web title", err);
        setWebTitle("Error loading title");
      }
    };

    fetchWebTitle();
  }, [sp]);

  return (
    <section className={`${styles.pnPContext} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
        <div>Connected Site Title: <strong>{escape(webTitle)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
          This web part is configured with <strong>SPFx 1.20</strong> and <strong>PnPjs</strong> using React Context.
        </p>
        <h4>Learn more about PnPjs:</h4>
        <ul className={styles.links}>
          <li><a href="https://pnp.github.io/pnpjs/" target="_blank" rel="noreferrer">PnPjs Documentation</a></li>
          <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
        </ul>
      </div>
    </section>
  );
};

export default PnPContext;
