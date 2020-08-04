import { IBaseSPListItem } from './Common';

export enum Status {
  Green,
  Yellow,
  Red,
  Grey
}

export class Project extends IBaseSPListItem {
  public Title: string;
  public ProjectNumber: number;
  public Client: string;
  public Unit: string;
  public ProjectManagerId: number;
  public SalesManagerId: number;
  public FinanceStatus: Status;
  public QualityStatus: Status;
  public ScopeStatus: Status;
  public TeamStatus: Status;
  public RelationStatus: Status;
  public TimeStatus: Status;
  public ProjectMembersId: Array<number>;
}