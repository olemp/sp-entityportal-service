import { Web, List, ItemAddResult } from '@pnp/sp';

export interface ISpEntityPortalServiceParams {
    webUrl: string,
    listName: string,
    groupIdFieldName: string,
    contentTypeId?: string,
    fieldsGroupName?: string,
}

export interface INewEntityResult {
    item: any;
    editFormUrl: string;
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
     * @param {string} groupId Group ID
     */
    public async getEntityItem(groupId: string): Promise<any> {
        try {
            const [item] = await this.list.items.filter(`${this.params.groupIdFieldName} eq '${groupId}'`).get();
            return item;
        } catch (e) {
            throw e;
        }
    }

    /**
     * Get entity item ID
     * 
     * @param {string} groupId Group ID
     */
    public async getEntityItemId(groupId: string): Promise<number> {
        try {
            const item = await this.getEntityItem(groupId);
            return item.Id;
        } catch (e) {
            throw e;
        }
    }

    /**
     * Get entity item field values
     * 
     * @param {string} groupId Group ID
     */
    public async getEntityItemFieldValues(groupId: string): Promise<any> {
        try {
            const itemId = await this.getEntityItemId(groupId);
            const itemFieldValues = await this.list.items.getById(itemId).fieldValuesAsText.get();
            return itemFieldValues;
        } catch (e) {
            throw e;
        }
    }

     /**
     * Get entity edit form url
     * 
     * @param {string} groupId Group ID
     * @param {string} sourceUrl Source URL
     * @param {number} _itemId Item id
     */
    public async getEntityEditFormUrl(groupId: string, sourceUrl: string, _itemId?: number): Promise<string> {
        try {
            const [itemId, { DefaultEditFormUrl }] = await Promise.all([
                _itemId ? (async () => _itemId)() : this.getEntityItemId(groupId),
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
     * @param {string} groupId Group ID
     * @param {Object} properties Properties
     */
    public async updateEntityItem(groupId: string, properties: { [key: string]: string }): Promise<any> {
        try {
            const itemId = await this.getEntityItemId(groupId);
            await this.list.items.getById(itemId).update(properties);
        } catch (e) {
            throw e;
        }
    }

    /**
     * New entity
     * 
     * @param {string} title Title
     * @param {string} groupId Group ID
     * @param {string} sourceUrl Source URL
     */
    public async newEntity(title: string, groupId: string, sourceUrl: string = null): Promise<INewEntityResult> {
        try {
            let properties = { Title: title };
            properties[this.params.groupIdFieldName] = groupId;
            const { data } = await this.list.items.add(properties);
            const editFormUrl = await this.getEntityEditFormUrl(groupId, sourceUrl, data.Id);
            return { item: data, editFormUrl };
        } catch (e) {
            throw e;
        }
    }
}
