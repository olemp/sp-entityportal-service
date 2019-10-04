import { IEntityField } from './IEntityField';
import { IEntityUrls } from './IEntityUrls';
import { TypedHash } from '@pnp/common';
export interface IEntity {
    /**
     * Item
     */
    item: TypedHash<any>;
    /**
     * Fields
     */
    fields: IEntityField[];
    /**
     * Urls
     */
    urls: IEntityUrls;
    /**
     * Field values
     */
    fieldValues: TypedHash<string>;
}
