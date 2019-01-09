import { Web, List, ItemAddResult } from '@pnp/sp';

export default class SpEntityPortalService {
    public web: Web;
    public list: List;

    constructor(
        public webUrl: string,
        public listName: string,
        public groupIdFieldName: string,
    ) {
        this.webUrl = webUrl;
        this.listName = listName;
        this.groupIdFieldName = groupIdFieldName;
        this.web = new Web(this.webUrl);
        this.list = this.web.lists.getByTitle(this.listName);
    }

    public async GetEntityItem(groupId: string): Promise<any> {
        try {
            const [item] = await this.list.items.filter(`${this.groupIdFieldName} eq '${groupId}'`).get();
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
            properties[this.groupIdFieldName] = groupId;
            return await this.list.items.add(properties);
        } catch (e) {
            throw e;
        }
    }
}
