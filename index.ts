import { stringIsNullOrEmpty } from '@pnp/core'
import '@pnp/sp/presets/all'
import {
  IContentType,
  IItem,
  IItemAddResult,
  IItemUpdateResult,
  IList,
  IWeb,
  SPFI,
  SPFx,
  spfi
} from '@pnp/sp/presets/all'
import {
  IEntity,
  IEntityField,
  IEntityUrls,
  INewEntityPermissions,
  ISpEntityPortalServiceParams
} from './types'

export class SpEntityPortalService {
  public sp: SPFI
  public web: IWeb
  private _entityList: IList
  private _entityContentType: IContentType

  /**
   * Construct a new `SpEntityPortalService` instance
   *
   * @param _spfxContext - SPFx context
   * @param _params - Parameters
   */
  constructor(private _spfxContext: any, private _params: ISpEntityPortalServiceParams) {
    this.sp = spfi(_params.portalUrl).using(SPFx(_spfxContext))
    this.web = this.sp.web
    this._entityList = this.web.lists.getByTitle(this._params.listName)
    this._entityContentType = this._params.contentTypeId
      ? this.sp.web.contentTypes.getById(this._params.contentTypeId)
      : null
  }

  /**
   * Returns a new instance of the SpEntityPortalService using the specified params
   *
   * @param params - Params
   */
  public usingParams(params: Partial<ISpEntityPortalServiceParams>) {
    return new SpEntityPortalService(this._spfxContext, { ...this._params, ...params })
  }

  /**
   * Get entity item
   *
   * @param identity - Identity
   * @param sourceUrl - Source URL used to generate URLs
   */
  public async fetchEntity(identity: string, sourceUrl: string): Promise<IEntity> {
    let [item, fields] = await Promise.all([this.getEntityItem(identity), this.getEntityFields()])
    let [urls, fieldValues] = await Promise.all([
      this.getEntityUrls(item.Id, sourceUrl),
      this.getEntityItemFieldValues(item.Id)
    ])
    return { item, fields, urls, fieldValues }
  }

  /**
   * Get entity fields
   */
  public async getEntityFields(): Promise<IEntityField[]> {
    if (!this._entityContentType) {
      return []
    }
    try {
      let query = this._entityContentType.fields.select(
        'Id',
        'InternalName',
        'Title',
        'Description',
        'TypeAsString',
        'SchemaXml',
        'TextField',
        'Group'
      )
      if (!stringIsNullOrEmpty(this._params.fieldPrefix)) {
        query = query.filter(`substringof('${this._params.fieldPrefix}', InternalName)`)
      }
      return await query<IEntityField[]>()
    } catch (e) {
      return []
    }
  }

  /**
   * Get entity item
   *
   * @param identity - Identity
   */
  public async getEntityItem<T = any>(identity: string): Promise<T> {
    try {
      if (identity.length === 38) {
        identity = identity.substring(1, 37)
      }
      return (
        await this._entityList.items.filter(`${this._params.identityFieldName} eq '${identity}'`)()
      )[0]
    } catch (e) {
      throw e
    }
  }

  /**
   * Get entity item field values
   *
   * @param itemId - Item ID
   */
  protected async getEntityItemFieldValues(itemId: number): Promise<{ [key: string]: any }> {
    try {
      return await this._entityList.items.getById(itemId).fieldValuesAsText()
    } catch (e) {
      throw e
    }
  }

  /**
   * Get entity urls
   *
   * @param itemId Item id
   * @param sourceUrl Source URL
   */
  protected async getEntityUrls(itemId: number, sourceUrl: string): Promise<IEntityUrls> {
    try {
      const { Id, DefaultEditFormUrl } = await this._entityList
        .select('DefaultEditFormUrl', 'Id')
        .expand('DefaultEditFormUrl')<{ Id: string; DefaultEditFormUrl: string }>()
      let editFormUrl = `${window.location.protocol}//${window.location.hostname}${DefaultEditFormUrl}?ID=${itemId}`
      let versionHistoryUrl = `${this._params.portalUrl}/_layouts/15/versions.aspx?list=${Id}&ID=${itemId}`
      if (sourceUrl) {
        editFormUrl += `&Source=${encodeURIComponent(sourceUrl)}`
        versionHistoryUrl = `&Source=${encodeURIComponent(sourceUrl)}`
      }
      return { editFormUrl, versionHistoryUrl }
    } catch (e) {
      throw e
    }
  }

  /**
   * Update enity item
   *
   * @param identity Identity
   * @param properties Properties
   */
  public async updateEntityItem(
    identity: string,
    properties: Record<string, string>
  ): Promise<IItemUpdateResult> {
    try {
      const item = await this.getEntityItem(identity)
      return await this._entityList.items.getById(item.Id).update(properties)
    } catch (e) {
      throw e
    }
  }

  /**
   * Create new entity
   *
   * @param identity Identity
   * @param url Url
   * @param additionalProperties Additional properties
   * @param permissions Permissions
   */
  public async createNewEntity(
    identity: string,
    url: string,
    additionalProperties?: { [key: string]: any },
    permissions?: INewEntityPermissions
  ): Promise<IItemAddResult> {
    try {
      let properties = { [this._params.identityFieldName]: identity, ...additionalProperties }
      if (this._params.urlFieldName) {
        properties[this._params.urlFieldName] = url
      }
      let itemAddResult = await this._entityList.items.add(properties)
      if (permissions) {
        await this.setEntityPermissions(itemAddResult.item, permissions)
      }
      return itemAddResult
    } catch (e) {
      throw e
    }
  }

  /**
   * Ensure entity exists with the specified `identity`. If it doesn't exist, it will be created.
   *
   * @param identity Identity
   * @param url URL of the entity
   * @param properties Properties
   * @returns
   */
  public async ensureEntity(
    identity: string,
    url: string,
    properties: Record<string, string>
  ): Promise<IItemAddResult | IItemUpdateResult> {
    try {
      let item = await this.getEntityItem(identity)
      if (!item) {
        return await this.createNewEntity(identity, url, properties)
      }
      return await this._entityList.items.getById(item.Id).update(properties)
    } catch (e) {
      throw e
    }
  }

  /**
   * Set entity permissions
   *
   * @param item Item/entity
   * @param permissions Permissions
   */
  private async setEntityPermissions(
    item: IItem,
    { fullControlPrincipals, readPrincipals, addEveryoneRead }: INewEntityPermissions
  ) {
    await item.breakRoleInheritance(false, true)
    if (fullControlPrincipals) {
      for (let i = 0; i < fullControlPrincipals.length; i++) {
        let principal = await this.sp.web.ensureUser(fullControlPrincipals[i])
        await item.roleAssignments.add(principal.data.Id, 1073741829)
      }
    }
    if (readPrincipals) {
      for (let i = 0; i < readPrincipals.length; i++) {
        let principal = await this.sp.web.ensureUser(readPrincipals[i])
        await item.roleAssignments.add(principal.data.Id, 1073741826)
      }
    }
    if (addEveryoneRead) {
      const [everyonePrincipal] = await this.sp.web.siteUsers
        .filter(`substringof('spo-grid-all-user', LoginName)`)
        .select('Id')<{ Id: number }[]>()
      await item.roleAssignments.add(everyonePrincipal.Id, 1073741826)
    }
  }
}

export { IEntity, IEntityField, IEntityUrls, ISpEntityPortalServiceParams }
