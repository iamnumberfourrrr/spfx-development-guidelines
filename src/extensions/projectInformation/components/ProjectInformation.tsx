import * as React from "react";
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { BaseReactiveComponent } from "../../../common/components/baseComponents/BaseReactiveComponent";
import { Project } from "../../../common/model/Project";
import { NavService } from "../../../common/services/NavService";
import { ProjectService } from "../../../common/services/ProjectService";
import { ServiceFactory } from "../../../common/services/ServiceFactory";
import styles from "./ProjectInformation.module.scss";
import ProjectLeaders from "./ProjectLeaders";
import ProjectMetadata from "./ProjectMetadata";
import ProjectStatuses from "./ProjectStatuses";

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
    const prjSubscription = this._projectSvc.currentProject$.subscribe(currentProject => currentProject && this.setState({ currentProject }));
    this.subs.push(prjSubscription);
  }

  public render() {
    const { currentProject } = this.state;

    return (
      <div className={styles.projectInformation}>
        {currentProject ? (
          <>
            <div className={styles.leftSection}>
              <ProjectMetadata project={currentProject} />
            </div>
            <div className={styles.rightSection}>
              <ProjectLeaders project={currentProject} context={this.props.context}/>
              <ProjectStatuses project={currentProject} />
            </div>
          </>
        ) : (
          <Spinner />
        )}
      </div>
    );
  }
}
