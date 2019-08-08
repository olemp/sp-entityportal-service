import { Web, List, Fields } from '@pnp/sp';
import { ISpEntityPortalServiceParams } from './ISpEntityPortalServiceParams';
import { INewEntityResult } from './INewEntityResult';
import { INewEntityPermissions } from './INewEntityPermissions';
import { IEntityField } from './IEntityField';
export default class SpEntityPortalService {
    params: ISpEntityPortalServiceParams;
    web: Web;
    list: List;
    contentType: any;
    fields: Fields;
    constructor(params: ISpEntityPortalServiceParams);
    /**
     * Get entity item fields
     */
    getEntityFields(): Promise<IEntityField[]>;
    /**
     * Get entity item
     *
     * @param {string} identity Identity
     */
    getEntityItem(identity: string): Promise<{
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
     */
    getEntityItemFieldValues(identity: string): Promise<{
        [key: string]: any;
    }>;
    /**
    * Get entity edit form url
    *
    * @param {string} identity Identity
    * @param {string} sourceUrl Source URL
    * @param {number} _itemId Item id
    */
    getEntityEditFormUrl(identity: string, sourceUrl: string, _itemId?: number): Promise<string>;
    /**
    * Get entity version history url
    *
    * @param {string} identity Identity
    * @param {string} sourceUrl Source URL
    * @param {number} _itemId Item id
    */
    getEntityVersionHistoryUrl(identity: string, sourceUrl: string, _itemId?: number): Promise<string>;
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
