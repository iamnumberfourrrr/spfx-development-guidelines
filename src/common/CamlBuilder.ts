import * as CalmJs from 'camljs';
import { IBaseSPListItem } from './model/Common';

class FinalizableToString {
    constructor(protected _exp: CalmJs.IFinalizableToString) { }
    /** Get the resulting CAML query as string */
    public ToString() {
        return this._exp.ToString();
    }
}

class Finalizable extends FinalizableToString {
    constructor(_exp: CalmJs.IFinalizable) {
        super(_exp);
    }
    public ToCamlQuery() {
        return (this._exp as CalmJs.IFinalizable).ToCamlQuery();
    }
}

class SortedQuery<T> extends Finalizable {
    constructor(_exp: CalmJs.ISortedQuery) {
        super(_exp);
    }
    public ThenBy(fieldInternalName: keyof T) {
        return (this._exp as CalmJs.ISortedQuery).ThenBy(fieldInternalName as string);
    }
    public ThenByDesc(fieldInternalName: keyof T) {
        return (this._exp as CalmJs.ISortedQuery).ThenByDesc(fieldInternalName as string);
    }
}

class Sortable<T> extends Finalizable {
    constructor(exp: CalmJs.ISortable) {
        super(exp);
    }
    public OrderBy(fieldInternalName: keyof T, override?: boolean, useIndexForOrderBy?: boolean): SortedQuery<T> {
        return new SortedQuery<T>((this._exp as CalmJs.ISortable).OrderBy(fieldInternalName as string, override, useIndexForOrderBy));
    }
    public OrderByDesc(fieldInternalName: keyof T, override?: boolean, useIndexForOrderBy?: boolean): SortedQuery<T> {
        return new SortedQuery<T>((this._exp as CalmJs.ISortable).OrderByDesc(fieldInternalName as string, override, useIndexForOrderBy));
    }
}

class Groupable<T> extends Sortable<T>{
    constructor(exp: CalmJs.IGroupable) {
        super(exp);
    }

    public GroupBy(fieldInternalName: keyof T, collapse?: boolean, groupLimit?: number): Sortable<T> {
        return new Sortable<T>((this._exp as CalmJs.IGroupable).GroupBy(fieldInternalName as string, collapse, groupLimit));
    }
}

class Query<T> extends Groupable<T> {
    constructor(exp: CalmJs.IQuery) {
        super(exp);
    }
    public Where() {
        // tslint:disable-next-line: no-use-before-declare
        return new FieldExpression<T>((this._exp as CalmJs.IQuery).Where());
    }
}

export class Expression<T> extends Groupable<T>{
    constructor(exp: CalmJs.IExpression) {
        super(exp);
    }

    public get IntExp() {
        return this._exp as CalmJs.IExpression;
    }

    public And(): FieldExpression<T> {
        // tslint:disable-next-line: no-use-before-declare
        return new FieldExpression<T>((this._exp as CalmJs.IExpression).And());
    }

    public Or(): FieldExpression<T> {
        // tslint:disable-next-line: no-use-before-declare
        return new FieldExpression<T>((this._exp as CalmJs.IExpression).Or());
    }
}

class TextFieldExpression<T> {
    constructor(private _exp: CalmJs.ITextFieldExpression) { }

    public EqualTo(value: string) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.EqualTo(value));
    }
    public NotEqualTo(value: string) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.NotEqualTo(value));
    }
    public Contains(value: string) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.Contains(value));
    }
    public BeginsWith(value: string) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.BeginsWith(value));
    }
    public IsNull() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsNull());
    }
    public IsNotNull() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsNotNull());
    }
    public In(arrayOfValues: string[]) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.In(arrayOfValues));
    }
}

class BooleanFieldExpression<T> {
    constructor(private _exp: CalmJs.IBooleanFieldExpression) { }

    public IsTrue() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsTrue());
    }
    public IsFalse() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsFalse());
    }
    public EqualTo(value: boolean) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.EqualTo(value));
    }
    public NotEqualTo(value: boolean) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.NotEqualTo(value));
    }
    public IsNull() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsNull());
    }
    public IsNotNull() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsNotNull());
    }
}

class NumberFieldExpression<T> {
    constructor(private _exp: CalmJs.INumberFieldExpression) { }
    public EqualTo(value: number) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.EqualTo(value));
    }
    public NotEqualTo(value: number) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.NotEqualTo(value));
    }
    public GreaterThan(value: number) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.GreaterThan(value));
    }
    public LessThan(value: number) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.LessThan(value));
    }
    public GreaterThanOrEqualTo(value: number) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.GreaterThanOrEqualTo(value));
    }
    public LessThanOrEqualTo(value: number) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.LessThanOrEqualTo(value));
    }
    public IsNull() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsNull());
    }
    public IsNotNull() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsNotNull());
    }
    public In(arrayOfValues: number[]) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.In(arrayOfValues));
    }
}

class UserFieldExpression<T> {
    constructor(private _exp: CalmJs.IUserFieldExpression) { }
    public EqualToCurrentUser() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.EqualToCurrentUser());
    }
    public IsInCurrentUserGroups() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsInCurrentUserGroups());
    }
    public IsInSPGroup(groupId: number) {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsInSPGroup(groupId));
    }
    public IsInSPWebGroups() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsInSPWebGroups());
    }
    public IsInSPWebAllUsers() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsInSPWebAllUsers());
    }
    public IsInSPWebUsers() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsInSPWebUsers());
    }
    public Id() {
        return new NumberFieldExpression<T>(this._exp.Id());
    }
    public ValueAsText() {
        return new TextFieldExpression<T>(this._exp.ValueAsText());
    }
}

class UserMultiFieldExpression<T> {
    constructor(private _exp: CalmJs.IUserMultiFieldExpression) { }

    public IncludesSuchItemThat() {
        return new UserFieldExpression<T>(this._exp.IncludesSuchItemThat());
    }

    public IsNull() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsNull());
    }

    public IsNotNull() {
        // tslint:disable-next-line: no-use-before-declare
        return new Expression<T>(this._exp.IsNotNull());
    }
    /** DEPRECATED: use "IncludesSuchItemThat().ValueAsText().EqualTo(value)" instead. */
    public Includes(value: any) {
        return new Expression<T>(this._exp.Includes(value));
    }
    /** DEPRECATED: use "IncludesSuchItemThat().ValueAsText().NotEqualTo(value)" instead. */
    public NotIncludes(value: any) {
        return new Expression<T>(this._exp.NotIncludes(value));
    }
}

class LookupFieldExpression<T> {
    constructor(private _exp: CalmJs.ILookupFieldExpression) { }
    public Id() {
        return new NumberFieldExpression<T>(this._exp.Id());
    }
    public ValueAsText() {
        return new TextFieldExpression<T>(this._exp.ValueAsText());
    }
    public ValueAsNumber() {
        return new NumberFieldExpression<T>(this._exp.ValueAsNumber());
    }
    public ValueAsDate() {
        // tslint:disable-next-line: no-use-before-declare
        return new DateTimeFieldExpression<T>(this._exp.ValueAsDate());
    }
    public ValueAsDateTime() {
        // tslint:disable-next-line: no-use-before-declare
        return new DateTimeFieldExpression<T>(this._exp.ValueAsDateTime());
    }
    public ValueAsBoolean() {
        return new BooleanFieldExpression<T>(this._exp.ValueAsBoolean());
    }
}

class LookupMultiFieldExpression<T> {
    constructor(private _exp: CalmJs.ILookupMultiFieldExpression) { }
    /** Checks a condition against every item in the multi lookup value */
    public IncludesSuchItemThat() {
        return new LookupFieldExpression<T>(this._exp.IncludesSuchItemThat());
    }
    /** Checks whether the field values collection is empty */
    public IsNull() {
        return new Expression(this._exp.IsNull());
    }
    /** Checks whether the field values collection is not empty */
    public IsNotNull() {
        return new Expression(this._exp.IsNotNull());
    }
    /** DEPRECATED: use "IncludesSuchItemThat().ValueAsText().EqualTo(value)" instead. */
    // Includes(value: any): IExpression;
    // /** DEPRECATED: use "IncludesSuchItemThat().ValueAsText().NotEqualTo(value)" instead. */
    // NotIncludes(value: any): IExpression;
    // /** DEPRECATED: "Eq" operation in CAML works exactly the same as "Includes". To avoid confusion, please use Includes. */
    // EqualTo(value: any): IExpression;
    // /** DEPRECATED: "Neq" operation in CAML works exactly the same as "NotIncludes". To avoid confusion, please use NotIncludes. */
    // NotEqualTo(value: any): IExpression;
}

class DateTimeFieldExpression<T> {
    constructor(private _exp: CalmJs.IDateTimeFieldExpression) { }
    public IsNull() {
        return new Expression<T>(this._exp.IsNull());
    }
    public IsNotNull() {
        return new Expression<T>(this._exp.IsNotNull());
    }
    public EqualTo(value: Date) {
        return new Expression<T>(this._exp.EqualTo(value));
    }
    public NotEqualTo(value: Date) {
        return new Expression<T>(this._exp.NotEqualTo(value));
    }
    /** Checks whether the value of the field is greater than the specified value */
    public GreaterThan(value: Date) {
        return new Expression(this._exp.GreaterThan(value));
    }
    /** Checks whether the value of the field is less than the specified value */
    public LessThan(value: Date) {
        return new Expression(this._exp.LessThan(value));
    }
    /** Checks whether the value of the field is greater than or equal to the specified value */
    public GreaterThanOrEqualTo(value: Date) {
        return new Expression(this._exp.GreaterThanOrEqualTo(value));
    }
    /** Checks whether the value of the field is less than or equal to the specified value */
    public LessThanOrEqualTo(value: Date) {
        return new Expression(this._exp.LessThanOrEqualTo(value));
    }
    /** Checks whether the value of the field is equal to one of the specified values */
    public In(arrayOfValues: Date[]) {
        return new Expression(this._exp.In(arrayOfValues));
    }
}

export class FieldExpression<T extends IBaseSPListItem> {
    constructor(private _epx: CalmJs.IFieldExpression) { }

    public All(...conditions: Expression<T>[]): Expression<T> {
        return new Expression<T>(this._epx.All(conditions.map(c => c.IntExp)));
    }

    public Any(...conditions: Expression<T>[]): Expression<T> {
        return new Expression<T>(this._epx.Any(conditions.map(c => c.IntExp)));
    }

    public TextField(internalName: keyof T): TextFieldExpression<T> {
        return new TextFieldExpression<T>(this._epx.TextField(internalName as string));
    }
    public BooleanField(internalName: keyof T): BooleanFieldExpression<T> {
        return new BooleanFieldExpression<T>(this._epx.BooleanField(internalName as string));
    }
    public UrlField(internalName: keyof T): TextFieldExpression<T> {
        return new TextFieldExpression<T>(this._epx.TextField(internalName as string));
    }
    public NumberField(internalName: keyof T): NumberFieldExpression<T> {
        return new NumberFieldExpression<T>(this._epx.NumberField(internalName as string));
    }
    public IntegerField(internalName: keyof T): NumberFieldExpression<T> {
        return new NumberFieldExpression<T>(this._epx.NumberField(internalName as string));
    }

    public UserField(internalName: keyof T): UserFieldExpression<T> {
        return new UserFieldExpression<T>(this._epx.UserField(internalName as string));
    }
    public LookupField(internalName: keyof T) {
        return new LookupFieldExpression<T>(this._epx.LookupField(internalName as string));
    }
    public LookupMultiField(internalName: keyof T) {
        return new LookupMultiFieldExpression<T>(this._epx.LookupMultiField(internalName as string));
    }
    public UserMultiField(internalName: keyof T): UserMultiFieldExpression<T> {
        return new UserMultiFieldExpression<T>(this._epx.UserMultiField(internalName as string));
    }
    public DateField(internalName: keyof T) {
        return new DateTimeFieldExpression<T>(this._epx.DateField(internalName as string));
    }
    public DateTimeField(internalName: keyof T) {
        return new DateTimeFieldExpression<T>(this._epx.DateField(internalName as string));
    }
}

export class CamlBuilder<T extends IBaseSPListItem> {
    private _builder = new CalmJs();

    public Where() {
        return new FieldExpression<T>(this._builder.Where());
    }

    public View(...viewFields: (keyof T)[]) {
        // tslint:disable-next-line: no-use-before-declare
        return new View<T>(this._builder.View(viewFields as string[]));
    }

    public static Expression<T>() {
        return new FieldExpression<T>(CalmJs.Expression());
    }
}

class Joinable<T, T1> {
    constructor(protected _exp: CalmJs.IJoinable) { }
    public InnerJoin(lookupFieldInternalName: keyof T, alias: string, fromList?: string): Join<T1> {
        // tslint:disable-next-line: no-use-before-declare
        return new Join<T1>(this._exp.InnerJoin(lookupFieldInternalName as string, alias, fromList));
    }
    public LeftJoin(lookupFieldInternalName: keyof T, alias: string, fromList?: string) {
        // tslint:disable-next-line: no-use-before-declare
        return new Join<T1>(this._exp.LeftJoin(lookupFieldInternalName as string, alias, fromList));
    }
}

class ProjectableView<T> extends Joinable<T, {}> {
    constructor(exp: CalmJs.IProjectableView) {
        super(exp);
    }

    public Query() {
        return new Query<T>((this._exp as CalmJs.IProjectableView).Query());
    }

    public RowLimit(limit: number, paged?: boolean): View<T> {
        // tslint:disable-next-line: no-use-before-declare
        return new View<T>((this._exp as CalmJs.IProjectableView).RowLimit(limit, paged));
    }

    public Scope(scope: CalmJs.ViewScope) {
        // tslint:disable-next-line: no-use-before-declare
        return new View<T>((this._exp as CalmJs.IProjectableView).Scope(scope));
    }

    public Select(remoteFieldInternalName: keyof T, remoteFieldAlias: string): ProjectableView<T> {
        return new ProjectableView<T>((this._exp as CalmJs.IProjectableView).Select(remoteFieldInternalName as string, remoteFieldAlias));
    }
}

class Join<T, T1 = {}> extends Joinable<T, T1> {
    constructor(exp: CalmJs.IJoin) {
        super(exp);
    }
    public Select(remoteFieldInternalName: keyof T, remoteFieldAlias: string) {
        return new ProjectableView<T>((this._exp as CalmJs.IJoin).Select(remoteFieldInternalName as string, remoteFieldAlias));
    }
}

class View<T> extends Finalizable {
    constructor(exp: CalmJs.IView) {
        super(exp);
    }
    public Query() {
        return new Query<T>((this._exp as CalmJs.IView).Query());
    }
    public RowLimit(limit: number, paged?: boolean): View<T> {
        return new View<T>((this._exp as CalmJs.IView).RowLimit(limit, paged));
    }
    public Scope(scope: CalmJs.ViewScope): View<T> {
        return new View<T>((this._exp as CalmJs.IView).Scope(scope));
    }
    public InnerJoin(lookupFieldInternalName: keyof T, alias: string): Join<T> {
        return new Join<T>((this._exp as CalmJs.IView).InnerJoin(lookupFieldInternalName as string, alias));
    }
    public LeftJoin(lookupFieldInternalName: keyof T, alias: string) {
        return new Join<T>((this._exp as CalmJs.IView).LeftJoin(lookupFieldInternalName as string, alias));
    }
}