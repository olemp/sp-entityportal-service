import { ItemAddResult, ItemUpdateResult, Web, SPConfiguration } from '@pnp/sp';
import { IEntity } from './IEntity';
import { IEntityField } from './IEntityField';
import { IEntityItem } from './IEntityItem';
import { IEntityUrls } from './IEntityUrls';
import { INewEntityPermissions } from './INewEntityPermissions';
import { ISpEntityPortalServiceParams } from './ISpEntityPortalServiceParams';
import { TypedHash } from '@pnp/common';
export declare class SpEntityPortalService {
    private _params;
    web: Web;
    private _entityList;
    private _entityContentType;
    /**
     * Constructor
     *
     * @param {ISpEntityPortalServiceParams} _params Parameters
     */
    constructor(_params: ISpEntityPortalServiceParams);
    /**
     * Configure
     *
     * @param {SPConfiguration} spConfiguration SP configuration
     */
    configure(spConfiguration?: SPConfiguration): SpEntityPortalService;
    /**
     * Returns a new instance of the SpEntityPortalService using the specified params
     *
     * @param {ISpEntityPortalServiceParams} params Params
     */
    usingParams(params: ISpEntityPortalServiceParams): SpEntityPortalService;
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
    updateEntityItem(identity: string, properties: TypedHash<string>): Promise<ItemUpdateResult>;
    /**
     * Create new entity
     *
     * @param {string} identity Identity
     * @param {string} url Url
     * @param {Object} additionalProperties Additional properties
     * @param {string} sourceUrl Source URL
     * @param {INewEntityPermissions} permissions Permissions
     */
    createNewEntity(identity: string, url: string, additionalProperties?: {
        [key: string]: any;
    }, permissions?: INewEntityPermissions): Promise<ItemAddResult>;
    /**
     * Set entity permissions
     *
     * @param {Item} item Item/entity
     * @param {INewEntityPermissions} permissions Permissions
     */
    private setEntityPermissions;
}
export { ISpEntityPortalServiceParams, IEntityField, IEntityItem, IEntity, IEntityUrls };
