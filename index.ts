import { Web, List, Item, Fields } from '@pnp/sp';
import { PageContext } from "@microsoft/sp-page-context";

export interface ISpEntityPortalServiceParams {
    webUrl: string;
    listName: string;
    identityFieldName: string;
    urlFieldName?: string;
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

export interface IEntityField {
    Title: string;
    InternalName: string;
    TypeAsString: string;
    SchemaXml: string;
}

export default class SpEntityPortalService {
    public web: Web;
    public list: List;
    public contentType: any;
    public fields: Fields;

    constructor(public params: ISpEntityPortalServiceParams) {
        this.params = params;
        this.web = new Web(this.params.webUrl);
        this.list = this.web.lists.getByTitle(this.params.listName);
        if (this.params.contentTypeId && this.params.fieldsGroupName) {
            this.contentType = this.web.contentTypes.getById(this.params.contentTypeId);
            this.fields = this.contentType.fields.filter(`Group eq '${this.params.fieldsGroupName}'`);
        }
    }

    /**
     * Get entity item fields
     */
    public async getEntityFields(): Promise<IEntityField[]> {
        if (!this.fields) {
            return null;
        }
        try {
            return this.fields.select('InternalName', 'Title', 'TypeAsString', 'SchemaXml').get<IEntityField[]>();
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
            const [item] = await this.list.items.filter(`${this.params.identityFieldName} eq '${identity}'`).get();
            if (item) {
                return item;
            } else {
                throw `Found no enity item with site ID ${identity}`;
            }
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
            const itemFieldValues = await this.list.items.getById(itemId).fieldValuesAsText.get();
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
    * @param {number} _itemId Item id
    */
    public async getEntityEditFormUrl(identity: string, sourceUrl: string, _itemId?: number): Promise<string> {
        try {
            const [itemId, { DefaultEditFormUrl }] = await Promise.all([
                _itemId ? (async () => _itemId)() : this.getEntityItemId(identity),
                this.list.select('DefaultEditFormUrl').expand('DefaultEditFormUrl').get(),
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
     * Update enity item
     * 
     * @param {any} context Context
     * @param {Object} properties Properties
     */
    public async updateEntityItem(context: any, properties: { [key: string]: string }): Promise<void> {
        try {
            const identity = (context as PageContext).site.id.toString();
            const itemId = await this.getEntityItemId(identity);
            await this.list.items.getById(itemId).update(properties);
        } catch (e) {
            throw e;
        }
    }

    /**
     * New entity
     * 
     * @param {string} identity Identity
     * @param {string} url Url
     * @param {string} sourceUrl Source URL
     * @param {INewEntityPermissions} permissions Permissions
     */
    public async newEntity(identity: string, url: string, sourceUrl: string = null, permissions?: INewEntityPermissions): Promise<INewEntityResult> {
        try {
            let properties = { Title: '' };
            properties[this.params.identityFieldName] = identity;
            if (this.params.urlFieldName) {
                properties[this.params.urlFieldName] = url;
            }
            const { data, item } = await this.list.items.add(properties);
            if (permissions) {
                await this.setEntityPermissions(item, permissions);
            }
            const editFormUrl = await this.getEntityEditFormUrl(identity, sourceUrl, data.Id);
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
                let principal = await this.web.ensureUser(fullControlPrincipals[i]);
                await item.roleAssignments.add(principal.data.Id, 1073741829);
            }
        }
        if (readPrincipals) {
            for (let i = 0; i < readPrincipals.length; i++) {
                let principal = await this.web.ensureUser(readPrincipals[i]);
                await item.roleAssignments.add(principal.data.Id, 1073741826);
            }
        }
        if (addEveryoneRead) {
            const [everyonePrincipal] = await this.web.siteUsers.filter(`substringof('spo-grid-all-user', LoginName)`).select('Id').get<Array<{ Id: number }>>();
            await item.roleAssignments.add(everyonePrincipal.Id, 1073741826);
        }
    }
}
