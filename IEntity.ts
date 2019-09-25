import { IEntityField } from './IEntityField';
import { IEntityItem } from './IEntityItem';
import { IEntityUrls } from './IEntityUrls';

export interface IEntity {
    /**
     * @todo Describe property
     */
    item: IEntityItem;

    /**
     * @todo Describe property
     */
    fields: IEntityField[];

    /**
     * @todo Describe property
     */
    urls: IEntityUrls;

    /**
     * @todo Describe property
     */
    fieldValues: {
        [key: string]: string;
    };
}
