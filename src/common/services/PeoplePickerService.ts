import { MSGraphClient, SPHttpClient } from "@microsoft/sp-http";
import { uniq } from '@microsoft/sp-lodash-subset';
import { IPersona, IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import * as Rx from 'rx-lite';
import { BaseListService } from "./BaseListService";

export interface IAdGroup {
    id: string;
    displayName: string;
    description: string;
  }

export enum PrincipleType {
    User = 1,
    SPGroup = 8,
    SecurityGroup = 4,
    DistributionGroup = 2,
    All = 15
}

export interface IUserEntityData {
    IsAltSecIdPresent: string;
    ObjectId: string;
    Title: string;
    Email: string;
    MobilePhone: string;
    OtherMails: string;
    Department: string;
}

export interface IClientPeoplePickerSearchUser {
    Key: string;
    Description: string;
    DisplayText: string;
    EntityType: string;
    ProviderDisplayName: string;
    ProviderName: string;
    IsResolved: boolean;
    EntityData: IUserEntityData;
    MultipleMatches: any[];
}

export interface IEnsureUser {
    Email: string;
    Id: number;
    IsEmailAuthenticationGuestUser: boolean;
    IsHiddenInUI: boolean;
    IsShareByEmailGuestUser: boolean;
    IsSiteAdmin: boolean;
    LoginName: string;
    PrincipalType: number;
    Title: string;
    UserId: {
        NameId: string;
        NameIdIssuer: string;
    };
}

export interface IEnsurableSharePointUser
    extends IClientPeoplePickerSearchUser, IEnsureUser { }

export class SharePointUserPersona implements IPersona {
    private _user: IEnsurableSharePointUser | IEnsureUser;
    public get User(): IEnsurableSharePointUser | IEnsureUser {
        return this._user;
    }

    public set User(user: IEnsurableSharePointUser | IEnsureUser) {
        this._user = user;
        this.primaryText = user.Title;
        const ensureSPUser = user as IEnsurableSharePointUser;
        let key = '';
        if (ensureSPUser.EntityData) {
            this.secondaryText = ensureSPUser.EntityData.Title || ensureSPUser.DisplayText;
            this.tertiaryText = ensureSPUser.EntityData.Department || ensureSPUser.Description;
            key = ensureSPUser.Key;
        } else {
            key = user.LoginName;
        }

        this.imageShouldFadeIn = true;

        this.imageUrl = `/_layouts/15/userphoto.aspx?size=S&accountname=${key.substr(key.lastIndexOf('|') + 1)}`;
    }

    constructor(user: IEnsurableSharePointUser | IEnsureUser) {
        this.User = user;
    }

    public primaryText: string;
    public secondaryText: string;
    public tertiaryText: string;
    public imageUrl: string;
    public imageShouldFadeIn: boolean;
}

export class UserService extends BaseListService {
    public Key = 'PeoplePickerService';
    private _resolvedUserByIds: Rx.BehaviorSubject<IEnsureUser[]> = new Rx.BehaviorSubject([]);
    private _resolvedGroupByIds: Rx.BehaviorSubject<IEnsureUser[]> = new Rx.BehaviorSubject([]);
    private _loadedUserIds: number[] = [];
    private _loadedGroupIds: number[] = [];
    private _ensuredUsers: { [key: string]: IEnsureUser } = {};
    private _currentUser: IEnsureUser;
    private _graphClient: MSGraphClient;

    public async getCurrentUserGroups(): Promise<IEnsureUser[]> {
        const resp = await this.context.spHttpClient.get(`${this.currentWebUrl}/_api/Web/CurrentUser/Groups`,
            SPHttpClient.configurations.v1, { headers: this.getGETHeaders() });
      
        return (await resp.json()).value as IEnsureUser[];
    }

    public async getGroupUsers(groupName: string): Promise<IEnsureUser[]> {
        const resp = await this.context.spHttpClient.get(`${this.currentWebUrl}/_api/web/sitegroups/getbyname('${groupName}')/users`,
            SPHttpClient.configurations.v1, { headers: this.getGETHeaders() });    
        return (await resp.json()).value as IEnsureUser[];
    }

    public async getCurrentUserAdGroups(): Promise<IAdGroup[]> {
        let resultData: IAdGroup[] = [];
        this._graphClient = await this.context.msGraphClientFactory.getClient();
        var result = await this._graphClient.api('me/memberOf').version('v1.0').get();
        if (result.error){
            console.error(result.error);
            return;
        }
        if (result.value.length > 0) {
            result.value.map((item: any) => {
                resultData.push({
                    id: item.id,
                    displayName: item.displayName,
                    description: item.description
                });
            });
        }
        return resultData;
    }
    
    public async isUserBelongToGroups(...groupName:string[]) {
        const currengUserAdGroups = await this.getCurrentUserAdGroups();
        return this.isUserBelogToGroup(groupName, currengUserAdGroups);
    }

    private async isUserBelogToGroup(groupNames:string[], userAdGroups:IAdGroup[]) {
        return new Promise<boolean>(async resolve => {
            const groupMembers = await this.getGroupUsers(groupNames.pop());
            const found = groupMembers.some(u => u.PrincipalType === PrincipleType.User && u.Email.toLowerCase() === this.context.pageContext.user.email.toLowerCase())
            || userAdGroups.some(g => groupMembers.some(u => u.Title.toLowerCase() === g.displayName.toLowerCase()));
            if (found) {
                resolve(true);
            } else if (groupNames.length>0) {
                resolve(await this.isUserBelogToGroup(groupNames, userAdGroups));
            } else {
                resolve(false);
            }
        });        
    }

    public async getCurrentUser(): Promise<IEnsureUser> {
        if (this._currentUser) {
            return this._currentUser;
        }

        this._currentUser = await this.get<IEnsureUser>(`${this.currentWebUrl}/_api/Web/CurrentUser`, true);
        return this._currentUser;
    }

    public async searchPeople(filter?: string): Promise<IPersonaProps[]> {
        const url = `${this.currentWebUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
        const requestBody = JSON.stringify({
            'queryParams': {
                '__metadata': {
                    'type': 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters'
                },
                'AllowEmailAddresses': true,
                'AllowMultipleEntities': false,    //if we want this is a multiple people picker then set 'true'
                'AllUrlZones': false,
                'MaximumEntitySuggestions': 50,
                'PrincipalSource': 15,
                'PrincipalType': PrincipleType.User,
                'QueryString': filter || ''
            }
        });

        const resp = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1,
            {
                body: requestBody,
                headers: {
                    ...this.getOdataVersionHeader('3.0'),
                    ...this.getGETHeaders(true)
                }
            });
        const data = await resp.json();
        let userQueryResults: IClientPeoplePickerSearchUser[] = JSON.parse(data.d.ClientPeoplePickerSearchUser);
        let persons = userQueryResults.map(p => new SharePointUserPersona(p as IEnsurableSharePointUser));
        await this.ensureUsers(persons);
        return persons;
    }

    public async ensureUsers(users: SharePointUserPersona[]) {
        const batchPromises: Promise<IEnsureUser>[] = users.map(p => this.ensureUser((p.User as IEnsurableSharePointUser).Key));
        const resolvedPeoples = await Promise.all(batchPromises);
        resolvedPeoples.forEach(p => {
            let userPersona = users.filter(up => (up.User as IEnsurableSharePointUser).Key === p.LoginName)[0];
            if (userPersona && userPersona.User) {
                userPersona.User = { ...userPersona.User, ...p };
            }
        });
    }

    public async ensureUserByLogins(usersIds: string[]) {
        const batchPromises: Promise<IEnsureUser>[] = usersIds.map(p => this.ensureUser(p));
        const resolvedPeoples = await Promise.all(batchPromises);
        return resolvedPeoples;
    }

    public async ensureUser(logonName: string): Promise<IEnsureUser> {
        if (this._ensuredUsers[logonName]) {
            return this._ensuredUsers[logonName];
        }

        const ensureUrl = `${this.currentWebUrl}/_api/web/ensureUser`;
        const resp = await this.context.spHttpClient.post(ensureUrl, SPHttpClient.configurations.v1,
            {
                body: JSON.stringify({ logonName: logonName }),
                headers: {
                    ...this.getGETHeaders(false),
                    ...this.getOdataVersionHeader('3.0')
                }
            });
        const resolvedUser = await resp.json();
        this._ensuredUsers[logonName] = resolvedUser as IEnsureUser;
        return resolvedUser as IEnsureUser;
    }

    public async getUserById(id: number):Promise<IEnsureUser> {
        const ensureUrl = `${this.currentWebUrl}/_api/web/siteusers/getById(${id})`;
        let resp = await this.context.spHttpClient.get(ensureUrl, SPHttpClient.configurations.v1,
            {
                headers: {
                    ...this.getGETHeaders(false)
                }
            });
            let resolvedUser;
            resolvedUser = await resp.json();
           return (resolvedUser);
    }
    private async loadUserById(id: number) {
        if (this._loadedUserIds.indexOf(id) >= 0) {
            return;
        }

        this._loadedUserIds.push(id);
        const ensureUrl = `${this.currentWebUrl}/_api/web/siteusers/getById(${id})`;
        let resp = await this.context.spHttpClient.get(ensureUrl, SPHttpClient.configurations.v1,
            {
                headers: {
                    ...this.getGETHeaders(false)
                }
            });
        let resolvedUser;
        if (resp.status === 200) {
            resolvedUser = await resp.json();
            const values = this._resolvedUserByIds.getValue();
            values.push(resolvedUser);
            this._resolvedUserByIds.onNext(values);
        } else {
            this._loadedUserIds = this._loadedUserIds.filter(i => i !== id);
            if (this._loadedGroupIds.indexOf(id) >= 0) {
                return;
            }

            this._loadedGroupIds.push(id);
            // user not found, could be group
            resp = await this.context.spHttpClient.get(`${this.currentWebUrl}/_api/web/sitegroups/getById(${id})`, SPHttpClient.configurations.v1, {
                headers: {
                    ...this.getGETHeaders(false)
                }
            });
            if (resp.status === 200) {
                resolvedUser = await resp.json();
            } else {
                resolvedUser = { Id: id } as any as IEnsureUser;
            }

            const values = this._resolvedGroupByIds.getValue();
            values.push(resolvedUser);
            this._resolvedGroupByIds.onNext(values);

            const userValues = this._resolvedUserByIds.getValue();
            userValues.push({ Id: id } as any);
            this._resolvedUserByIds.onNext(userValues);
        }
    }

    public async getUserByIds(...ids: number[]): Promise<IEnsureUser[]> {
        const uniqueIds = uniq(ids);
        uniqueIds.forEach(i => this.loadUserById(i));
        return new Promise<IEnsureUser[]>(resolve => {
            this._resolvedUserByIds.map(users => users.filter(u => uniqueIds.indexOf(u.Id) >= 0))
                .filter(vals => vals.length === uniqueIds.length)
                .take(1)
                .subscribe((val) => {
                    const groupIds = uniq(val.filter(v => v.Id && !v.Title).map(v => v.Id));
                    if (groupIds.length === 0) {
                        resolve(val);
                    } else {
                        const users = val.filter(v => v.Id && v.Email);
                        this._resolvedGroupByIds.map(groups => groups.filter(g => groupIds.indexOf(g.Id) >= 0))
                            .filter(vals => vals.length === groupIds.length)
                            .take(1).subscribe(groupVal => {
                                resolve([...users, ...groupVal]);
                            });
                    }
                });
        });
    }

    public async getCurrentUserGroupMembership(groupName: string): Promise<boolean> {
        const resp = await this.context.spHttpClient.get(`${this.currentWebUrl}/_api/web/sitegroups/getbyname('${groupName}')/CanCurrentUserViewMembership`,
            SPHttpClient.configurations.v1, { headers: this.getGETHeaders() });
        var res = (await resp.json()).value as boolean;
        return res;
    }
}