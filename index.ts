import { Web, List } from '@pnp/sp';

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

    public async GetEntityItemId(groupId: string): Promise<any> {
        try {
            const item = await this.GetEntityItem(groupId);
            return item.Id;
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
}
