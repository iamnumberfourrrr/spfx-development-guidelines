import * as React from 'react';
import styles from './ProjectsList.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProjectsListProps {
  context: WebPartContext;
}


export default class ProjectsList extends React.Component<IProjectsListProps, {}> {
  public render(): React.ReactElement<IProjectsListProps> {
    return (
      <div className={ styles.projectsList }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
