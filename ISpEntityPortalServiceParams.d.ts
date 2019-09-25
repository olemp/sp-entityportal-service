export interface ISpEntityPortalServiceParams {
    /**
     * Portal URL
     */
    portalUrl: string;
    /**
     * List name for the entities
     */
    listName: string;
    /**
     * Field name that indentifies the entity
     */
    identityFieldName: string;
    /**
     * Field name for site url
     */
    urlFieldName?: string;
    /**
     * Content type ID for entity
     */
    contentTypeId?: string;
    /**
     * Field prefix for entity fields
     */
    fieldPrefix?: string;
}
