import { Web, List } from '@pnp/sp';
export default class SpEntityPortalService {
    webUrl: string;
    listName: string;
    groupIdFieldName: string;
    web: Web;
    list: List;
    constructor(webUrl: string, listName: string, groupIdFieldName: string);
    GetEntityItem(groupId: string): Promise<any>;
    GetEntityItemId(groupId: string): Promise<any>;
    UpdateEntityItem(groupId: string, properties: {
        [key: string]: string;
    }): Promise<any>;
}
