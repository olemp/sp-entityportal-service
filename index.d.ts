import { Web, List, ItemAddResult } from '@pnp/sp';
export interface ISpEntityPortalServiceParams {
    webUrl: string;
    listName: string;
    groupIdFieldName: string;
    contentTypeId?: string;
    fieldsGroupName?: string;
}
export default class SpEntityPortalService {
    params: ISpEntityPortalServiceParams;
    web: Web;
    list: List;
    contentType: any;
    fields: any;
    constructor(params: ISpEntityPortalServiceParams);
    GetEntityFields(): Promise<any[]>;
    GetEntityItem(groupId: string): Promise<any>;
    GetEntityItemId(groupId: string): Promise<number>;
    GetEntityItemFieldValues(groupId: string): Promise<any>;
    GetEntityEditFormUrl(groupId: string, sourceUrl: string): Promise<string>;
    UpdateEntityItem(groupId: string, properties: {
        [key: string]: string;
    }): Promise<any>;
    NewEntity(title: string, groupId: string): Promise<ItemAddResult>;
}
