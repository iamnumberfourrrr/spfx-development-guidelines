import { BaseComponentContext } from "@microsoft/sp-component-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as uuid from 'uuid';
import { IBaseSPListItem } from "../model/Common";
import { ODataQueryBuilder } from "../ODataQueryBuilder";
import { getGlobalCache } from "../Utils";
import { IOdataPagingInfo } from "../model/IOdataPagingInfo";

export const JSON_HEADER = 'application/json';
export const JSON_HEADER_VERBOSE = `${JSON_HEADER};odata=verbose`;

export const COMMON_HEADERS = {
    'accept': JSON_HEADER,
    'content-type': JSON_HEADER
};

export const COMMON_HEADERS_VERBOSE = {
    'accept': JSON_HEADER_VERBOSE,
    'content-type': JSON_HEADER_VERBOSE
};

export const EDIT_ITEM_HEADERS = {
    'IF-MATCH': '*',
    'X-HTTP-Method': 'MERGE'
};

export const DELETE_ITEM_HEADERS = {
    'IF-MATCH': '*',
    'X-HTTP-Method': 'DELETE'
};

export const LIST_REST_API = '/_api/web/lists';

export class BatchBodyBuilder {
    private _changesetId: string;
    private _batchGuid: string;
    private _changesetContents: string[] = [];
    constructor() {
        this._changesetId = uuid.v4();
        this._batchGuid = uuid.v4();
    }

    public addOperation(type: 'GET' | 'POST' | 'MERGE' | 'DELETE' | 'PUT', endpoint: string, headers: { [key: string]: string }, body: string) {
        this._changeSetSeparator();
        this._changesetContents.push(`Content-Type: application/http`);
        this._changesetContents.push('Content-Transfer-Encoding: binary');
        this._changesetContents.push('');
        this._changesetContents.push(`${type} ${endpoint} HTTP/1.1`);
        Object.keys(headers).forEach(key => {
            this._changesetContents.push(`${key}: ${headers[key]}`);
        });
        this._changesetContents.push('');
        this._changesetContents.push(body);
        this._changesetContents.push('');
    }

    public getBody() {
        this._changeSetSeparator();
        const changesetBody = this._changesetContents.join('\n');
        const batchContents = [];
        // create batch for update items
        batchContents.push('--batch_' + this._batchGuid);
        batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + this._changesetId + '"');
        batchContents.push('');
        batchContents.push(changesetBody);
        batchContents.push('');
        batchContents.push('--batch_' + this._batchGuid + '--');
        return batchContents.join('\n');
    }

    public getHeaders(additionHeader?: { [key: string]: string }) {
        if (!additionHeader) {
            additionHeader = {};
        }

        return {
            'accept': 'application/json',
            'Content-Type': 'multipart/mixed; boundary="batch_' + this._batchGuid + '"',
            'odata-version': ' ',
            ...additionHeader
        };
    }

    private _changeSetSeparator() {
        this._changesetContents.push(`--changeset_${this._changesetId}`);
    }
}

export abstract class BaseListService {
    public abstract Key: string;
    public currentWebUrl = '';
    public hostUrl = '';

    constructor(public context?: BaseComponentContext, serviceFactory = false) {        
        if (!serviceFactory){
            this.init();
        }
    }

    public init(){
        if (this.context) {
            this.currentWebUrl = this.context.pageContext.web.absoluteUrl;
            this.hostUrl = this.context.pageContext.site.absoluteUrl.replace(this.context.pageContext.site.serverRelativeUrl, '');
        }
        this.onInit();
    }

    protected onInit(){

    }

    protected expandField(field: string, subField: string) {
        return `${field}/${subField}`;
    }

    protected escapeStringValue(val: string) {
        return (val || "").replace(/'/g, "''");
    }

    protected getOdataVersionHeader(version?: string) {
        return { 'odata-version': version };
    }

    protected getGETHeaders(verbose?: boolean) {
        return verbose ? COMMON_HEADERS_VERBOSE : COMMON_HEADERS;
    }

    protected getEditItemHeaders() {
        return {
            ...this.getOdataVersionHeader(' '), // must specify this other will get error Parsing JSON Light feeds or entries in requests without entity set is not supported
            ...COMMON_HEADERS_VERBOSE,
            ...EDIT_ITEM_HEADERS
        };
    }

    protected getListItemsByListGuidRESTURL(webUrl: string, listGuid: string, itemId: string) {
        return `${webUrl}${LIST_REST_API}(guid'${listGuid}')/items(${itemId})`;
    }


    protected getListItemsByListTitleRESTURL(webUrl: string, listTitle: string, id?: number | string) {
        return `${this.getListByListTitleRESTURL(webUrl, listTitle)}/items${id ? ('(' + id + ')') : ''}`;
    }

    protected getListByListTitleRESTURL(webUrl: string, listTitle: string) {
        return `${webUrl}${LIST_REST_API}/getbytitle('${listTitle}')`;
    }

    protected getListUrl(webUrl: string, listName: string) {
        return `${webUrl}/Lists/${listName}`;
    }

    protected escapseListFilterText(text: string) {
        return (text || "").replace(/'/g, '');
    }

    protected async getListItemEntityTypeFullName(webUrl: string, listTitle: string): Promise<string> {
        const cache = getGlobalCache('ListItemEntityTypeFullName');
        const listUrlApi = this.getListByListTitleRESTURL(webUrl, listTitle);
        if (cache[listUrlApi]) {
            return cache[listUrlApi];
        }

        const typeFullName = await this.get<string>(`${listUrlApi}/ListItemEntityTypeFullName`);
        cache[listUrlApi] = typeFullName;
        return typeFullName;
    }

    protected async addItem<T extends IBaseSPListItem>(webUrl: string, listTitle: string, values: T): Promise<T> {
        const itemTypeFullName = await this.getListItemEntityTypeFullName(webUrl, listTitle);
        let postUrl = this.getListItemsByListTitleRESTURL(webUrl, listTitle);
        const resp = await this.context.spHttpClient.post(postUrl, SPHttpClient.configurations.v1,
            {
                body: this.getEditJSONBody(itemTypeFullName, values),
                headers: {
                    ...this.getGETHeaders(true),
                    ...this.getOdataVersionHeader('3.0')
                }
            });
        return (await this.getJsonOrThrowIfError<{ d: T }>(resp)).d;
    }

    public async deleteItem(webUrl: string, listTitle: string, id: string | number) {
        const url = this.getListItemsByListTitleRESTURL(webUrl, listTitle, id);
        return await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1,
            {
                headers: {
                    ...this.getGETHeaders(false),
                    ...DELETE_ITEM_HEADERS
                }
            });
    }

    protected async editItem<T extends IBaseSPListItem>(webUrl: string, listTitle: string, values: T, itemId: number | string): Promise<T> {
        const itemTypeFullName = await this.getListItemEntityTypeFullName(webUrl, listTitle);
        let postUrl = this.getListItemsByListTitleRESTURL(webUrl, listTitle) + `(${itemId})`;
        await this.context.spHttpClient.post(postUrl, SPHttpClient.configurations.v1,
            {
                body: this.getEditJSONBody<T>(itemTypeFullName, values),
                headers: this.getEditItemHeaders()
            });
        return values as T;
    }

    protected getEditJSONBody<T extends IBaseSPListItem>(itemType: string, values: T): string {
        return JSON.stringify({
            '__metadata': { 'type': itemType },
            ...values as any
        });
    }

    protected async getJsonOrThrowIfError<T>(resp: SPHttpClientResponse): Promise<T> {
        const result = await resp.json();
        if (result.error) {
            throw result.error;
        }
        return result as T;
    }

    public async getListItemByCaml<T>(webUrl: string, listTitle: string, whereCond: string, viewFields?: string[]): Promise<T[]> {
        if (whereCond.indexOf('<Query>') < 0) {
            whereCond = `<Query>${whereCond}</Query>`;
        }

        if (viewFields && viewFields.length) {
            let viewFieldsXml = '';
            viewFields.forEach(field => viewFieldsXml = viewFieldsXml + `<FieldRef Name='${field}' />`);
            viewFieldsXml = `<ViewFields>${viewFieldsXml}</ViewFields>`;
            whereCond = `${whereCond}${viewFieldsXml}`;

        }

        if (whereCond.indexOf('<View>') < 0) {
            whereCond = `<View>${whereCond}</View>`;
        }

        const postBody = {
            headers: {
                ...this.getOdataVersionHeader('3.0')
            },
            body: JSON.stringify({
                "query": {
                    "__metadata": { "type": "SP.CamlQuery" },
                    "ViewXml": whereCond
                }
            })
        };
        const url = this.getListByListTitleRESTURL(webUrl, listTitle) + '/GetItems';
        const resp = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, postBody);
        return (await resp.json()).value as T[];
    }

    protected async getListChoiceFieldValues(listTitle: string, fieldName: string): Promise<string[]> {
        const query = new ODataQueryBuilder()
            .filter(f => f.eq('EntityPropertyName', fieldName))
            .select('Choices').toQuery();
        const url = this.getListByListTitleRESTURL(this.currentWebUrl, listTitle) + `/fields${query}`;
        const result = await this.get<{ Choices: string[] }[]>(url);
        return result[0].Choices;
    }

    protected async getListItems<T>(webUrl: string, listTitle: string, query = ''){
        const url = this.getListItemsByListTitleRESTURL(webUrl, listTitle) + query;
        return this.get<T[]>(url);
    }

    protected async getListItemsWithPaging<T>(webUrl: string, listTitle: string, query = '') {
        const url = this.getListItemsByListTitleRESTURL(webUrl, listTitle) + query;
        return this.getWithPageInfo<T[]>(url);
    }
    protected async getListItemsWithPagingURL<T>(webUrl: string) {
        return this.getWithPageInfo<T[]>(webUrl);
    }

    protected async get<T>(url: string, unwrapped = false, headers?: HeadersInit) {
        const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1,
            {
                headers: headers || this.getGETHeaders(false)
            });
        if (unwrapped) {
            return (await this.getJsonOrThrowIfError<T>(resp));
        }
        return (await this.getJsonOrThrowIfError<{ value: T }>(resp)).value;
    }


    protected async getWithPageInfo<T>(url: string, headers?: HeadersInit) {
        const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1,
            {
                headers: headers || this.getGETHeaders(false)
            });
      
            const json = (await this.getJsonOrThrowIfError<{ value: T }>(resp));
        const value = json.value;
        const nextLink = json["@odata.nextLink"];
        const pageInfo = new IOdataPagingInfo<T>();
        pageInfo.NextLink = nextLink;
        pageInfo.Value = value;

        return pageInfo;

    }
}