import { Web, List, ItemAddResult } from '@pnp/sp';

export interface ISpEntityPortalServiceParams {
    webUrl: string,
    listName: string,
    groupIdFieldName: string,
    contentTypeId?: string,
    fieldsGroupName?: string,
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

    public async GetEntityFields(): Promise<any[]> {
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

    public async GetEntityItem(groupId: string): Promise<any> {
        try {
            const [item] = await this.list.items.filter(`${this.params.groupIdFieldName} eq '${groupId}'`).get();
            return item;
        } catch (e) {
            throw e;
        }
    }

    public async GetEntityItemId(groupId: string): Promise<number> {
        try {
            const item = await this.GetEntityItem(groupId);
            return item.Id;
        } catch (e) {
            throw e;
        }
    }

    public async GetEntityItemFieldValues(groupId: string): Promise<any> {
        try {
            const itemId = await this.GetEntityItemId(groupId);
            const itemFieldValues = await this.list.items.getById(itemId).fieldValuesAsText.get();
            return itemFieldValues;
        } catch (e) {
            throw e;
        }
    }

    public async GetEntityEditFormUrl(groupId: string, sourceUrl: string): Promise<string> {
        try {
            const [itemId, { DefaultEditFormUrl }] = await Promise.all([
                this.GetEntityItemId(groupId),
                this.list.select('DefaultEditFormUrl').expand('DefaultEditFormUrl').get(),
            ]);
            return `${window.location.protocol}//${window.location.hostname}${DefaultEditFormUrl}?ID=${itemId}&Source=${encodeURIComponent(sourceUrl)}`;
        } catch (e) {
            throw e;
        }
    }

    public async UpdateEntityItem(groupId: string, properties: { [key: string]: string }): Promise<any> {
        try {
            const itemId = await this.GetEntityItemId(groupId);
            await this.list.items.getById(itemId).update(properties);
        } catch (e) {
            throw e;
        }
    }

    public async NewEntity(title: string, groupId: string): Promise<ItemAddResult> {
        try {
            let properties = { Title: title };
            properties[this.params.groupIdFieldName] = groupId;
            return await this.list.items.add(properties);
        } catch (e) {
            throw e;
        }
    }
}
