import { Web, List, ItemAddResult } from '@pnp/sp';
export default class SpEntityPortalService {
    webUrl: string;
    listName: string;
    groupIdFieldName: string;
    contentTypeName: string;
    fieldsGroupName: string;
    web: Web;
    list: List;
    contentType: any;
    fields: any;
    constructor(webUrl: string, listName: string, groupIdFieldName: string, contentTypeName: string, fieldsGroupName: string);
    GetEntityFields(): Promise<any[]>;
    GetEntityItem(groupId: string): Promise<any>;
    GetEntityItemId(groupId: string): Promise<number>;
    GetEntityItemFieldValues(groupId: string): Promise<number>;
    GetEntityEditFormUrl(groupId: string, sourceUrl: string): Promise<string>;
    UpdateEntityItem(groupId: string, properties: {
        [key: string]: string;
    }): Promise<any>;
    NewEntity(title: string, groupId: string): Promise<ItemAddResult>;
}
