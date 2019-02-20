import { Web, List } from '@pnp/sp';
export interface ISpEntityPortalServiceParams {
    webUrl: string;
    listName: string;
    siteIdFieldName: string;
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
     * @param {string} siteId Site ID
     */
    getEntityItem(context: any): Promise<any>;
    /**
     * Get entity item ID
     *
     * @param {any} context Context
     */
    getEntityItemId(context: any): Promise<number>;
    /**
     * Get entity item field values
     *
     * @param {any} context Context
     */
    getEntityItemFieldValues(context: any): Promise<any>;
    /**
    * Get entity edit form url
    *
     * @param {any} context Context
    * @param {string} sourceUrl Source URL
    * @param {number} _itemId Item id
    */
    getEntityEditFormUrl(context: any, sourceUrl: string, _itemId?: number): Promise<string>;
    /**
     * Update enity item
     *
     * @param {string} siteId Site ID
     * @param {Object} properties Properties
     */
    updateEntityItem(context: any, properties: {
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
