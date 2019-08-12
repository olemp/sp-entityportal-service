import { sp, Web, List, ContentType, Item, Fields } from '@pnp/sp';
import { ISpEntityPortalServiceParams } from './ISpEntityPortalServiceParams';
import { INewEntityResult } from './INewEntityResult';
import { INewEntityPermissions } from './INewEntityPermissions';
import { IEntityField } from './IEntityField';

sp.setup({ defaultCachingStore: "session", defaultCachingTimeoutSeconds: 60, globalCacheDisable: false });


export default class SpEntityPortalService {
    private _web: Web;
    private _list: List;
    private _contentType: ContentType;
    private _fields: IEntityField[];
    private _item: any;

    constructor(private params: ISpEntityPortalServiceParams) {
        this._web = new Web(this.params.webUrl);
        this._list = this._web.lists.getByTitle(this.params.listName);
        if (this.params.contentTypeId && this.params.fieldsGroupName) {
            this._contentType = this._web.contentTypes.getById(this.params.contentTypeId);
        }
    }

    /**
     * Get entity fields
     */
    public async getEntityFields(): Promise<IEntityField[]> {
        if (!this._contentType) {
            return null;
        }
        try {
            if (this._fields) return this._fields;
            this._fields = await this._contentType.fields
                .select('InternalName', 'Title', 'TypeAsString', 'SchemaXml')
                .filter(`Group eq '${this.params.fieldsGroupName}'`)
                .usingCaching()
                .get<IEntityField[]>();
            return this._fields;
        } catch (e) {
            throw e;
        }
    }


    /**
     * Get entity item
     * 
     * @param {string} identity Identity
     */
    public async getEntityItem(identity: string): Promise<{ [key: string]: any }> {
        try {
            if (identity.length === 38) {
                identity = identity.substring(1, 37);
            }
            if (this._item) return this._item;
            this._item = (await this._list.items.filter(`${this.params.identityFieldName} eq '${identity}'`).usingCaching().get())[0];
            return this._item;
        } catch (e) {
            throw e;
        }
    }

    /**
     * Get entity item ID
     * 
     * @param {string} identity Identity
     */
    public async getEntityItemId(identity: string): Promise<number> {
        try {
            const item = await this.getEntityItem(identity);
            return item.Id;
        } catch (e) {
            throw e;
        }
    }

    /**
     * Get entity item field values
     * 
     * @param {string} identity Identity
     */
    public async getEntityItemFieldValues(identity: string): Promise<{ [key: string]: any }> {
        try {
            const itemId = await this.getEntityItemId(identity);
            const itemFieldValues = await this._list.items.getById(itemId).fieldValuesAsText.usingCaching().get();
            return itemFieldValues;
        } catch (e) {
            throw e;
        }
    }

    /**
    * Get entity edit form url
    * 
    * @param {string} identity Identity
    * @param {string} sourceUrl Source URL
    */
    public async getEntityEditFormUrl(identity: string, sourceUrl: string): Promise<string> {
        try {
            const [itemId, { DefaultEditFormUrl }] = await Promise.all([
                this.getEntityItemId(identity),
                this._web.lists.getByTitle(this.params.listName).select('DefaultEditFormUrl').expand('DefaultEditFormUrl').usingCaching().get(),
            ]);
            let editFormUrl = `${window.location.protocol}//${window.location.hostname}${DefaultEditFormUrl}?ID=${itemId}`;
            if (sourceUrl) {
                editFormUrl += `&Source=${encodeURIComponent(sourceUrl)}`;
            }
            return editFormUrl;
        } catch (e) {
            throw e;
        }
    }

    /**
    * Get entity version history url
    * 
    * @param {string} identity Identity
    * @param {string} sourceUrl Source URL
    */
    public async getEntityVersionHistoryUrl(identity: string, sourceUrl: string): Promise<string> {
        try {
            const [itemId, { Id }] = await Promise.all([
                this.getEntityItemId(identity),
                this._web.lists.getByTitle(this.params.listName).select('Id').usingCaching().get(),
            ]);
            let editFormUrl = `${this.params.webUrl}/_layouts/15/versions.aspx?_list=${Id}&ID=${itemId}`;
            if (sourceUrl) {
                editFormUrl += `&Source=${encodeURIComponent(sourceUrl)}`;
            }
            return editFormUrl;
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
            const itemId = await this.getEntityItemId(identity);
            await this._list.items.getById(itemId).update(properties);
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
            const editFormUrl = await this.getEntityEditFormUrl(identity, sourceUrl);
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

export { ISpEntityPortalServiceParams };