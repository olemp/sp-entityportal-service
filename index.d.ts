import { Web, List } from '@pnp/sp';
export interface ISpEntityPortalServiceParams {
    webUrl: string;
    listName: string;
    groupIdFieldName: string;
    siteUrlFieldName?: string;
    contentTypeId?: string;
    fieldsGroupName?: string;
}
export interface INewEntityResult {
    item: any;
    editFormUrl: string;
}
export interface INewEntityPermissions {
    fullControlPrincipals?: string[];
    readPrincipals?: string[];
    addEveryoneRead?: boolean;
}
export default class SpEntityPortalService {
    params: ISpEntityPortalServiceParams;
    web: Web;
    list: List;
    contentType: any;
    fields: any;
    constructor(params: ISpEntityPortalServiceParams);
    /**
     * Get entity item fields
     */
    getEntityFields(): Promise<any[]>;
    /**
     * Get entity item
     *
     * @param {string} groupId Group ID
     */
    getEntityItem(groupId: string): Promise<any>;
    /**
     * Get entity item ID
     *
     * @param {string} groupId Group ID
     */
    getEntityItemId(groupId: string): Promise<number>;
    /**
     * Get entity item field values
     *
     * @param {string} groupId Group ID
     */
    getEntityItemFieldValues(groupId: string): Promise<any>;
    /**
    * Get entity edit form url
    *
    * @param {string} groupId Group ID
    * @param {string} sourceUrl Source URL
    * @param {number} _itemId Item id
    */
    getEntityEditFormUrl(groupId: string, sourceUrl: string, _itemId?: number): Promise<string>;
    /**
     * Update enity item
     *
     * @param {string} groupId Group ID
     * @param {Object} properties Properties
     */
    updateEntityItem(groupId: string, properties: {
        [key: string]: string;
    }): Promise<any>;
    /**
     * New entity
     *
     * @param {any} context Context
     * @param {string} sourceUrl Source URL
     * @param {INewEntityPermissions} permissions Permissions
     */
    newEntity(context: any, sourceUrl?: string, permissions?: INewEntityPermissions): Promise<INewEntityResult>;
    /**
     * Set entity permissions
     *
     * @param {Item} item Item/entity
     * @param {INewEntityPermissions} permissions Permissions
     */
    private setEntityPermissions;
}
