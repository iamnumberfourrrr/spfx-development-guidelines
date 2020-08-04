import * as React from 'react';
import styles from './ProjectInformation.module.scss';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { Project } from '../../../common/model/Project';
import { NavService } from '../../../common/services/NavService';
import { IEnsureUser, UserService } from '../../../common/services/PeoplePickerService';
import { ServiceFactory } from '../../../common/services/ServiceFactory';

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
      <div className={styles.projectRoles}>
        <div>Project Manager</div>
        <div>Sales Manager</div>
      </div>
      <div className={styles.memberNames}>
        <div>{projectManager && projectManager.Title}</div>
        <div>{salesManager && salesManager.Title}</div>
      </div>
    </div>;
  }
}