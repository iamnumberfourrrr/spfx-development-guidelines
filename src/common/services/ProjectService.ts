import { ServiceFactory } from './ServiceFactory';
import { UserService, IEnsureUser } from './PeoplePickerService';
import { ODataQueryBuilder } from "./../ODataQueryBuilder";
import { CamlBuilder } from "./../CamlBuilder";
import { RootListConstants } from "./../CommonConstants";
import { NavService } from "./NavService";
import { Project } from "./../model/Project";
import { BaseListService } from "./BaseListService";
import * as Rx from 'rx-lite';

export class ProjectService extends BaseListService {
  public Key = "ProjectListService";
  public currentProject$ = new Rx.BehaviorSubject<Project>(undefined);
  public currentProjectNumber: string;

  private _userSvc: UserService;

  protected onInit() {
    this.currentProjectNumber = NavService.getProjectNumberFromUrl();
    this._userSvc = ServiceFactory.getService<UserService>(UserService, this.context, NavService.getPortalSiteUrl());
  }

  public async loadCurrentProjectWithCamlQuery() {
    const whereCond = new CamlBuilder<Project>()
      .Where()
      .NumberField("ProjectNumber")
      .EqualTo(parseInt(this.currentProjectNumber))
      .ToString();
    const projects = await this.getListItemByCaml<Project>(this.currentWebUrl, RootListConstants.ProjectsLists, whereCond);
    const project = projects && projects.length > 0 ? projects[0] : null;
    this.currentProject$.onNext(project);
  }

  public async loadCurrentProjectWithOdata() {
    const query = new ODataQueryBuilder<Project>()
      .filter((f) => f.eq("ProjectNumber", this.currentProjectNumber))
      .toQuery();
    const projects = await this.getListItems<Project>(this.currentWebUrl, RootListConstants.ProjectsLists, query);
    const project = projects && projects.length > 0 ? projects[0] : null;
    this.currentProject$.onNext(project);
  }
}
