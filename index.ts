import { Web, List, ContentType, Item, } from '@pnp/sp';
import { ISpEntityPortalServiceParams } from './ISpEntityPortalServiceParams';
import { INewEntityResult } from './INewEntityResult';
import { INewEntityPermissions } from './INewEntityPermissions';
import { IEntityField } from './IEntityField';
import { IEntityItem } from './IEntityItem';
import { IEntityUrls } from './IEntityUrls';
import { IEntity } from './IEntity';

export class SpEntityPortalService {
    private _web: Web;
    private _list: List;
    private _contentType: ContentType;

    constructor(private params: ISpEntityPortalServiceParams) {
        this._web = new Web(this.params.webUrl);
        this._list = this._web.lists.getByTitle(this.params.listName);
        if (this.params.contentTypeId && this.params.fieldsGroupName) {
            this._contentType = this._web.contentTypes.getById(this.params.contentTypeId);
        }
    }

    /**
     * Get entity item
     * 
     * @param {string} identity Identity
     * @param {string} sourceUrl Source URL used to generate URLs
     */
    public async fetchEntity(identity: string, sourceUrl: string): Promise<IEntity> {
        let [item, fields] = await Promise.all([
            this.getEntityItem(identity),
            this.getEntityFields(),
        ]);
        let [urls, fieldValues] = await Promise.all([
            this.getEntityUrls(item.Id, sourceUrl),
            this.getEntityItemFieldValues(item.Id),
        ]);
        return { item, fields, urls, fieldValues };
    }

    /**
     * Get entity fields 
     */
    protected async getEntityFields(): Promise<IEntityField[]> {
        if (!this._contentType) {
            return [];
        }
        try {
            return await this._contentType.fields
                .select('InternalName', 'Title', 'TypeAsString', 'SchemaXml')
                .filter(`Group eq '${this.params.fieldsGroupName}'`)
                .get<IEntityField[]>();
        } catch (e) {
            return [];
        }
    }


    /**
     * Get entity item
     * 
     * @param {string} identity Identity
     */
    protected async getEntityItem(identity: string): Promise<IEntityItem> {
        try {
            if (identity.length === 38) {
                identity = identity.substring(1, 37);
            }
            return (
                await this._list.items
                    .filter(`${this.params.identityFieldName} eq '${identity}'`)
                    .get()
            )[0];
        } catch (e) {
            throw e;
        }
    }

    /**
     * Get entity item field values
     * 
    * @param {number} itemId Item id
     */
    protected async getEntityItemFieldValues(itemId: number): Promise<{ [key: string]: any }> {
        try {
            return await this._list.items
                .getById(itemId)
                .fieldValuesAsText
                .get();
        } catch (e) {
            throw e;
        }
    }

    /**
    * Get entity urls
    * 
    * @param {number} itemId Item id
    * @param {string} sourceUrl Source URL
    */
    protected async getEntityUrls(itemId: number, sourceUrl: string): Promise<IEntityUrls> {
        try {
            const { Id, DefaultEditFormUrl } = await this._list
                .select('DefaultEditFormUrl', 'Id')
                .expand('DefaultEditFormUrl')
                .get<{ Id: string, DefaultEditFormUrl: string }>();
            let editFormUrl = `${window.location.protocol}//${window.location.hostname}${DefaultEditFormUrl}?ID=${itemId}`;
            let versionHistoryUrl = `${this.params.webUrl}/_layouts/15/versions.aspx?list=${Id}&ID=${itemId}`;
            if (sourceUrl) {
                editFormUrl += `&Source=${encodeURIComponent(sourceUrl)}`;
                versionHistoryUrl = `&Source=${encodeURIComponent(sourceUrl)}`;
            }
            return { editFormUrl, versionHistoryUrl };
        } catch (e) {
            throw e;
        }
    }

    /**
     * Update enity item
     * 
     * @param {string} identity Identity
     * @param {Object} properties Properties
     */
    public async updateEntityItem(identity: string, properties: { [key: string]: string }): Promise<void> {
        try {
            const item = await this.getEntityItem(identity);
            await this._list.items.getById(item.Id).update(properties);
        } catch (e) {
            throw e;
        }
    }

    /**
     * New entity
     * 
     * @param {string} identity Identity
     * @param {string} url Url
     * @param {Object} additionalProperties Additional properties
     * @param {string} sourceUrl Source URL
     * @param {INewEntityPermissions} permissions Permissions
     */
    public async newEntity(identity: string, url: string, additionalProperties?: { [key: string]: any }, sourceUrl: string = null, permissions?: INewEntityPermissions): Promise<INewEntityResult> {
        try {
            let properties = { [this.params.identityFieldName]: identity, ...additionalProperties };
            if (this.params.urlFieldName) {
                properties[this.params.urlFieldName] = url;
            }
            const { data, item } = await this._list.items.add(properties);
            if (permissions) {
                await this.setEntityPermissions(item, permissions);
            }
            const { editFormUrl } = await this.getEntityUrls(data.Id, sourceUrl);
            return { item: data, editFormUrl };
        } catch (e) {
            throw e;
        }
    }

    /**
     * Set entity permissions
     * 
     * @param {Item} item Item/entity
     * @param {INewEntityPermissions} permissions Permissions
     */
    private async setEntityPermissions(item: Item, { fullControlPrincipals, readPrincipals, addEveryoneRead }: INewEntityPermissions) {
        await item.breakRoleInheritance(false, true);
        if (fullControlPrincipals) {
            for (let i = 0; i < fullControlPrincipals.length; i++) {
                let principal = await this._web.ensureUser(fullControlPrincipals[i]);
                await item.roleAssignments.add(principal.data.Id, 1073741829);
            }
        }
        if (readPrincipals) {
            for (let i = 0; i < readPrincipals.length; i++) {
                let principal = await this._web.ensureUser(readPrincipals[i]);
                await item.roleAssignments.add(principal.data.Id, 1073741826);
            }
        }
        if (addEveryoneRead) {
            const [everyonePrincipal] = await this._web.siteUsers.filter(`substringof('spo-grid-all-user', LoginName)`).select('Id').get<Array<{ Id: number }>>();
            await item.roleAssignments.add(everyonePrincipal.Id, 1073741826);
        }
    }
}

export { ISpEntityPortalServiceParams, INewEntityResult, IEntityField, IEntityItem, IEntity, IEntityUrls };