import * as React from "react";
import ProjectMetadata from "./ProjectMetadata";
import ProjectStatuses from "./ProjectStatuses";
import ProjectLeaders from "./ProjectLeaders";
import styles from "./ProjectInformation.module.scss";
import { Project } from "../../../common/model/Project";
import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { ServiceFactory } from "../../../common/services/ServiceFactory";
import { ProjectService } from "../../../common/services/ProjectService";
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { NavService } from "../../../common/services/NavService";
import { BaseReactiveComponent } from "../../../common/components/BaseReactiveComponent";

interface IProjectInformationState {
  currentProject: Project;
}

export default class ProjectInformation extends BaseReactiveComponent<{ context: BaseComponentContext},IProjectInformationState> {

  private _projectSvc: ProjectService;

  constructor(props) {
    super(props);
    this.state = {
      currentProject: undefined,
    };
  }

  public componentDidMount() {
    this._projectSvc = ServiceFactory.getService<ProjectService>(ProjectService, this.props.context, NavService.getPortalSiteUrl());
    this._projectSvc.loadCurrentProjectWithOdata();
    const prjSubscription = this._projectSvc.currentProject$.subscribe(currentProject => {
      currentProject && this.setState({ currentProject });
    })
    this.subs.push(prjSubscription);
  }

  public render() {
    const { currentProject } = this.state;

    return (
      <div className={styles.projectInformation}>
        {currentProject ? (
          <>
            <ProjectMetadata project={currentProject} />
            <ProjectLeaders project={currentProject} context={this.props.context} />
            <ProjectStatuses />
          </>
        ) : (
          <Spinner />
        )}
      </div>
    );
  }
}
