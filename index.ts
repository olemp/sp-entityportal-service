import { Web, List, Item } from '@pnp/sp';
import { PageContext } from "@microsoft/sp-page-context";

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
    public web: Web;
    public list: List;
    public contentType: any;
    public fields: any;

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
    public async getEntityFields(): Promise<any[]> {
        if (!this.fields) {
            return null;
        }
        try {
            const fields = await this.fields.get();
            return fields;
        } catch (e) {
            throw e;
        }
    }


    /**
     * Get entity item
     * 
     * @param {string} siteId Site ID
     */
    public async getEntityItem(siteId: string): Promise<any> {
        try {
            const [item] = await this.list.items.filter(`${this.params.siteIdFieldName} eq '${siteId}'`).get();
            if (item) {
                return item;
            } else {
                throw `Found no enity item with site ID ${siteId}`;
            }
        } catch (e) {
            throw e;
        }
    }

    /**
     * Get entity item ID
     * 
     * @param {string} siteId Site ID
     */
    public async getEntityItemId(siteId: string): Promise<number> {
        try {
            const item = await this.getEntityItem(siteId);
            return item.Id;
        } catch (e) {
            throw e;
        }
    }

    /**
     * Get entity item field values
     * 
     * @param {string} siteId Site ID
     */
    public async getEntityItemFieldValues(siteId: string): Promise<any> {
        try {
            const itemId = await this.getEntityItemId(siteId);
            const itemFieldValues = await this.list.items.getById(itemId).fieldValuesAsText.get();
            return itemFieldValues;
        } catch (e) {
            throw e;
        }
    }

    /**
    * Get entity edit form url
    * 
    * @param {string} siteId Site ID
    * @param {string} sourceUrl Source URL
    * @param {number} _itemId Item id
    */
    public async getEntityEditFormUrl(siteId: string, sourceUrl: string, _itemId?: number): Promise<string> {
        try {
            const [itemId, { DefaultEditFormUrl }] = await Promise.all([
                _itemId ? (async () => _itemId)() : this.getEntityItemId(siteId),
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
     * @param {string} siteId Site ID
     * @param {Object} properties Properties
     */
    public async updateEntityItem(siteId: string, properties: { [key: string]: string }): Promise<any> {
        try {
            const itemId = await this.getEntityItemId(siteId);
            await this.list.items.getById(itemId).update(properties);
        } catch (e) {
            throw e;
        }
    }

    /**
     * New entity
     * 
     * @param {any} context Context
     * @param {string} sourceUrl Source URL
     * @param {INewEntityPermissions} permissions Permissions
     */
    public async newEntity(context: any, sourceUrl: string = null, permissions?: INewEntityPermissions): Promise<INewEntityResult> {
        try {
            let properties = { Title: context.web.title };
            properties[this.params.siteIdFieldName] = (context as PageContext).site.id.toString();
            if (this.params.siteUrlFieldName) {
                properties[this.params.siteUrlFieldName] = (context as PageContext).web.absoluteUrl;
            }
            const { data, item } = await this.list.items.add(properties);
            if (permissions) {
                await this.setEntityPermissions(item, permissions);
            }
            const editFormUrl = await this.getEntityEditFormUrl((context as PageContext).site.id.toString(), sourceUrl, data.Id);
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
