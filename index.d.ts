import { ISpEntityPortalServiceParams } from './ISpEntityPortalServiceParams';
import { INewEntityResult } from './INewEntityResult';
import { INewEntityPermissions } from './INewEntityPermissions';
import { IEntityField } from './IEntityField';
import { IEntityItem } from './IEntityItem';
import { IEntityUrls } from './IEntityUrls';
import { IEntity } from './IEntity';
export declare class SpEntityPortalService {
    private params;
    private _web;
    private _list;
    private _contentType;
    constructor(params: ISpEntityPortalServiceParams);
    /**
     * Get entity item
     *
     * @param {string} identity Identity
     * @param {string} sourceUrl Source URL used to generate URLs
     */
    fetchEntity(identity: string, sourceUrl: string): Promise<IEntity>;
    /**
     * Get entity fields
     */
    getEntityFields(): Promise<IEntityField[]>;
    /**
     * Get entity item
     *
     * @param {string} identity Identity
     */
    getEntityItem(identity: string): Promise<IEntityItem>;
    /**
     * Get entity item field values
     *
    * @param {number} itemId Item id
     */
    protected getEntityItemFieldValues(itemId: number): Promise<{
        [key: string]: any;
    }>;
    /**
    * Get entity urls
    *
    * @param {number} itemId Item id
    * @param {string} sourceUrl Source URL
    */
    protected getEntityUrls(itemId: number, sourceUrl: string): Promise<IEntityUrls>;
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
export { ISpEntityPortalServiceParams, INewEntityResult, IEntityField, IEntityItem, IEntity, IEntityUrls };
