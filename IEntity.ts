import { IEntityField } from './IEntityField';
import { IEntityItem } from './IEntityItem';
import { IEntityUrls } from './IEntityUrls';
export interface IEntity {
    item: IEntityItem;
    fields: IEntityField[];
    urls: IEntityUrls;
    fieldValues: {
        [key: string]: string;
    };
}
