import { IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";

export const YesNoOptions: IDropdownOption[] = [
    { key: 1, text: 'Yes' },
    { key: 0, text: 'No' }
];

export class IBaseSPListItem {
    public ID?: number;
    public Title?: string;
    public ContentTypeId?: string;
    public "odata.editLink"?: string;
    public "odata.type"?: string;
}

export function getFieldName<P>(key: keyof P) {
    return key;
}

export function getFieldNames<P>(...keys: Array<keyof P>) {
    return keys;
}

export type PostUserMultiType = { results: number[] };