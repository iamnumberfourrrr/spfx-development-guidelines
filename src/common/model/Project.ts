import { IBaseSPListItem } from './Common';

export enum Status {
  Green,
  Yellow,
  Red,
  Grey
}

export class Project extends IBaseSPListItem {
  Title: string;
  ProjectNumber: number;
  Client: string;
  Unit: string;
  ProjectManagerId: number;
  SalesManagerId: number;
  FinanceStatus: Status;
  QualityStatus: Status;
  ScopeStatus: Status;
  TeamStatus: Status;
  TimeStatus: Status;
  ProjectMembersId: Array<number>;
}