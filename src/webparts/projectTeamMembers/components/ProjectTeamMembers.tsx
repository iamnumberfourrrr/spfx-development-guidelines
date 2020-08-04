import * as React from 'react';
import styles from './ProjectTeamMembers.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ProjectService } from '../../../common/services/ProjectService';
import { UserService, IEnsureUser } from '../../../common/services/PeoplePickerService';
import { ServiceFactory } from '../../../common/services/ServiceFactory';
import { NavService } from '../../../common/services/NavService';
import { BaseReactiveComponent } from '../../../common/components/baseComponents/BaseReactiveComponent';
import TeamMember from './TeamMember';

export interface IProjectTeamMembersProps {
  context: WebPartContext;
}

interface IProjectTeamMembersState {
  users: IEnsureUser[];
}

export default class ProjectsList extends BaseReactiveComponent<IProjectTeamMembersProps, IProjectTeamMembersState> {

  private _projectSvc: ProjectService;
  private _userSvc: UserService;

  constructor(props) {
    super(props);

    this.state = {
      users: []
    }

    this._projectSvc = ServiceFactory.getService<ProjectService>(ProjectService, this.props.context, NavService.getPortalSiteUrl());
    this._userSvc = ServiceFactory.getService<UserService>(UserService, this.props.context, NavService.getPortalSiteUrl());
  }

  public componentDidMount() {
    const projectSubscription = this._projectSvc.currentProject$.subscribe(async currentProject => {
      if (currentProject) {
        const users = await this._userSvc.getUserByIds(...currentProject.ProjectMembersId);
        this.setState({users});
      }
    });

    this.subs.push(projectSubscription);
  }

  public render(): React.ReactElement<IProjectTeamMembersProps> {
    const { users } = this.state;
    return (
      <div className={styles.projectTeamMembers}>
        <h1>Team members</h1>
        {users.map(u => <TeamMember user={u} />)}
      </div>
    );
  }
}
