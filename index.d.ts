import { Web, List } from '@pnp/sp';
export interface ISpEntityPortalServiceParams {
    webUrl: string;
    listName: string;
    groupIdFieldName: string;
    contentTypeId?: string;
    fieldsGroupName?: string;
}
export interface INewEntityResult {
    item: any;
    editFormUrl: string;
}
export default class SpEntityPortalService {
    params: ISpEntityPortalServiceParams;
    web: Web;
    list: List;
    contentType: any;
    fields: any;
    constructor(params: ISpEntityPortalServiceParams);
    getEntityFields(): Promise<any[]>;
    getEntityItem(groupId: string): Promise<any>;
    getEntityItemId(groupId: string): Promise<number>;
    getEntityItemFieldValues(groupId: string): Promise<any>;
    getEntityEditFormUrl(groupId: string, sourceUrl: string, _itemId?: number): Promise<string>;
    updateEntityItem(groupId: string, properties: {
        [key: string]: string;
    }): Promise<any>;
    /**
     * New entity
     *
     * @param {string} title Title
     * @param {string} groupId Group ID
     */
    newEntity(title: string, groupId: string): Promise<INewEntityResult>;
}
