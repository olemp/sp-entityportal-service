import { ISpEntityPortalServiceParams } from './ISpEntityPortalServiceParams';
import { INewEntityResult } from './INewEntityResult';
import { INewEntityPermissions } from './INewEntityPermissions';
import { IEntityField } from './IEntityField';
export default class SpEntityPortalService {
    private params;
    private _web;
    private _list;
    private _contentType;
    constructor(params: ISpEntityPortalServiceParams);
    /**
     * Get entity fields
     *
     * @param {Date} expiration Expiration
     */
    getEntityFields(expiration?: Date): Promise<IEntityField[]>;
    /**
     * Get entity item
     *
     * @param {string} identity Identity
     * @param {Date} expiration Expiration
     */
    getEntityItem(identity: string, expiration?: Date): Promise<{
        [key: string]: any;
    }>;
    /**
     * Get entity item ID
     *
     * @param {string} identity Identity
     */
    getEntityItemId(identity: string): Promise<number>;
    /**
     * Get entity item field values
     *
     * @param {string} identity Identity
     * @param {Date} expiration Expiration
     */
    getEntityItemFieldValues(identity: string, expiration?: Date): Promise<{
        [key: string]: any;
    }>;
    /**
    * Get entity edit form url
    *
    * @param {string} identity Identity
    * @param {string} sourceUrl Source URL
     * @param {Date} expiration Expiration
    */
    getEntityEditFormUrl(identity: string, sourceUrl: string, expiration?: Date): Promise<string>;
    /**
    * Get entity version history url
    *
    * @param {string} identity Identity
    * @param {string} sourceUrl Source URL
     * @param {Date} expiration Expiration
    */
    getEntityVersionHistoryUrl(identity: string, sourceUrl: string, expiration?: Date): Promise<string>;
    /**
     * Update enity item
     *
     * @param {string} identity Identity
     * @param {Object} properties Properties
     */
    updateEntityItem(identity: string, properties: {
        [key: string]: string;
    }): Promise<void>;
    /**
     * New entity
     *
     * @param {string} identity Identity
     * @param {string} url Url
     * @param {Object} additionalProperties Additional properties
     * @param {string} sourceUrl Source URL
     * @param {INewEntityPermissions} permissions Permissions
     */
    newEntity(identity: string, url: string, additionalProperties?: {
        [key: string]: any;
    }, sourceUrl?: string, permissions?: INewEntityPermissions): Promise<INewEntityResult>;
    /**
     * Set entity permissions
     *
     * @param {Item} item Item/entity
     * @param {INewEntityPermissions} permissions Permissions
     */
    private setEntityPermissions;
}
export { ISpEntityPortalServiceParams };
