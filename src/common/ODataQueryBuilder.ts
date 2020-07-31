import FilterBuilder, { QueryBuilder } from 'odata-query-builder';

export enum ODataLogicalOperators {
    eq = 'eq',
    ne = 'ne',
    lt = 'lt',
    gt = 'gt',
    le = 'le',
    ge = 'ge'
}

type filterExpressionType = string | number | boolean | Date;

export class ODataFilterBuilder<T = any> {
    private _isGrouped = false;
    private _oldToQuery: Function;
    constructor(private _filterBuilder?: FilterBuilder) {
        const filter = this.filterBuilder;
    }

    public get filterBuilder() {
        if (!this._filterBuilder) {
            this._filterBuilder = new FilterBuilder();
            this._oldToQuery = this._filterBuilder.toQuery;
            this._filterBuilder.toQuery = this.toQuery.bind(this);
        }
        return this._filterBuilder;
    }

    public group() {
        this._isGrouped = true;
        return this;
    }

    public filterExpression(field: keyof T, operator: ODataLogicalOperators, value: filterExpressionType) {
        this._filterBuilder.filterExpression(field as string, operator, value);
        return this;
    }

    public eq(field: keyof T, value: filterExpressionType) {
        return this.filterExpression(field, ODataLogicalOperators.eq, value);
    }

    public lt(field: keyof T, value: filterExpressionType) {
        return this.filterExpression(field, ODataLogicalOperators.lt, value);
    }

    public gt(field: keyof T, value: filterExpressionType) {
        return this.filterExpression(field, ODataLogicalOperators.gt, value);
    }

    public ne(field: keyof T, value: filterExpressionType) {
        return this.filterExpression(field, ODataLogicalOperators.ne, value);
    }

    public le(field: keyof T, value: filterExpressionType) {
        return this.filterExpression(field, ODataLogicalOperators.le, value);
    }

    public ge(field: keyof T, value: filterExpressionType) {
        return this.filterExpression(field, ODataLogicalOperators.ge, value);
    }

    public isNull(field: keyof T) {
        this.filterBuilder.filterPhrase(`${field} ${ODataLogicalOperators.eq} null`);
        return this;
    }

    public isNotNull(field: keyof T) {
        this.filterBuilder.filterPhrase(`${field} ${ODataLogicalOperators.ne} null`);
        return this;
    }

    public startswith(field: keyof T, value: string) {
        this.filterBuilder.filterPhrase(`startswith(${field},'${value}')`);
        return this;
    }

    public filterPhrase(phrase: string) {
        this.filterBuilder.filterPhrase(phrase);
        return this;
    }

    public and(predicate: (filter: ODataFilterBuilder<T>) => ODataFilterBuilder<T>) {
        this.filterBuilder.and(f => predicate(new ODataFilterBuilder<T>(f)).filterBuilder);
        return this;
    }

    public or(predicate: (filter: ODataFilterBuilder<T>) => ODataFilterBuilder<T>) {
        this.filterBuilder.or(f => predicate(new ODataFilterBuilder<T>(f)).filterBuilder);
        return this;
    }

    public toQuery(operator: string) {
        if (this._isGrouped) {
            return `(${this._oldToQuery(operator)})`;
        }
        return this._oldToQuery(operator);
    }
}

/**
 * Extends QueryBuilder with Select
 */
export class ODataQueryBuilder<T = any> {
    constructor(private _url?: string) { }

    private _selectFields: string[] = [];
    private _query = new QueryBuilder();

    public selectExt(...fields: string[]) {
        this._selectFields = fields.slice() as string[];
        return this;
    }

    public select(...fields: (keyof T)[]) {
        this._selectFields = fields.slice() as string[];
        return this;
    }

    public orderBy(...fields: string[]) {
        this._query.orderBy(fields.join(','));
        return this;
    }

    public top(top: number) {
        this._query.top(top);
        return this;
    }

    public skip(skip: number) {
        this._query.skip(skip);
        return this;
    }

    public count() {
        this._query.count();
        return this;
    }

    public expand(...fields: string[]) {
        this._query.expand(fields.join(','));
        return this;
    }

    public filter(predicate: (filter: ODataFilterBuilder<T>) => ODataFilterBuilder<T>, operator?: 'and' | 'or') {
        this._query.filter(f => predicate(new ODataFilterBuilder(f)).filterBuilder, operator);
        return this;
    }

    public toQuery() {
        let url = (this._url || '') + this._query.toQuery();
        let selectClause = '';
        if (this._selectFields.length > 0) {
            selectClause = `$select=${this._selectFields.join(',')}`;
            if (url.indexOf('?') >= 0) {
                url += '&' + selectClause;
            } else {
                url += '?' + selectClause;
            }
        }
        return url;
    }
}