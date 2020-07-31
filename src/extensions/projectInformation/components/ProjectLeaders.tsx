import * as React from 'react';
import styles from './ProjectInformation.module.scss';
import { Project } from '../../../common/model/Project';
import { ProjectService } from '../../../common/services/ProjectService';
import { ServiceFactory } from '../../../common/services/ServiceFactory';
import { NavService } from '../../../common/services/NavService';
import { IEnsureUser, UserService } from '../../../common/services/PeoplePickerService';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface ProjectLeadersProps {
  project: Project;
  context: BaseComponentContext;
}

interface ProjectLeadersState {
  projectManager: IEnsureUser;
  salesManager: IEnsureUser;
}

export default class ProjectLeaders extends React.Component<ProjectLeadersProps, ProjectLeadersState> {

  private _userSvc: UserService;

  constructor(props) {
    super(props);
    this.state = {
      projectManager: undefined,
      salesManager: undefined
    };
  }

  public async componentDidMount() {
    const { project, context } = this.props;
    const { ProjectManagerId, SalesManagerId } = project;

    this._userSvc = ServiceFactory.getService<UserService>(UserService, context, NavService.getPortalSiteUrl());
    const users = await this._userSvc.getUserByIds(ProjectManagerId, SalesManagerId);
    this.setState({
      projectManager: users.find((u) => u.Id === ProjectManagerId),
      salesManager: users.find((u) => u.Id === SalesManagerId),
    });
  }
  
  public render() {
    const { projectManager, salesManager } = this.state;
    return <div className={styles.projectLeaders}>
      <div className={styles.projectRole}>
        <div className={styles.roleTitle}>Project Manager: </div>
        <div className={styles.personName}>{projectManager && projectManager.Title}</div>
      </div>
      <div className={styles.projectRole}>
        <div className={styles.roleTitle}>Sales Manager: </div>
        <div className={styles.personName}>{salesManager && salesManager.Title}</div>
      </div>
    </div>
  }
}