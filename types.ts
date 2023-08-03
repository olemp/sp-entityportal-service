export interface IEntity {
  /**
   * Item
   */
  item: Record<string, any>

  /**
   * Fields
   */
  fields: IEntityField[]

  /**
   * Urls
   */
  urls: IEntityUrls

  /**
   * Field values
   */
  fieldValues: Record<string, string>
}

export interface IEntityField {
  Id?: string
  Title?: string
  Description?: string
  InternalName?: string
  TypeAsString?: string
  TextField?: string
  SchemaXml?: string
  Group?: string
}

export interface IEntityUrls {
  /**
   * Edit form URL for the entity
   */
  editFormUrl: string

  /**
   * Version history URL for the entity
   */
  versionHistoryUrl: string
}

export interface INewEntityPermissions {
  /**
   * @todo Describe property
   */
  fullControlPrincipals?: string[]

  /**
   * @todo Describe property
   */
  readPrincipals?: string[]

  /**
   * @todo Describe property
   */
  addEveryoneRead?: boolean
}

export interface ISpEntityPortalServiceParams {
  /**
   * SPFx context
   */
  spfxContext: any

  /**
   * Portal URL
   */
  portalUrl: string

  /**
   * List name for the entities
   */
  listName: string

  /**
   * Field name that indentifies the entity
   */
  identityFieldName: string

  /**
   * Field name for site url
   */
  urlFieldName: string

  /**
   * Field prefix for entity fields (optional)
   */
  fieldPrefix?: string

  /**
   * Content type ID for entity (optional)
   */
  contentTypeId?: string
}
