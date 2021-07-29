import { ITypedHash, stringIsNullOrEmpty } from '@pnp/common'
import {
  IContentType,
  IItem,
  IItemAddResult,
  IItemUpdateResult,
  IList,
  ISPConfiguration,
  IWeb,
  sp,
  Web
} from '@pnp/sp/presets/all'
import {
  IEntity,
  IEntityField,
  IEntityUrls,
  INewEntityPermissions,
  ISpEntityPortalServiceParams
} from './types'

export class SpEntityPortalService {
  public web: IWeb
  private _entityList: IList
  private _entityContentType: IContentType

  /**
   * Construct a new `SpEntityPortalService` instance
   *
   * @param _params - Parameters
   */
  constructor(private _params: ISpEntityPortalServiceParams) {
    this.web = Web(this._params.portalUrl)
    this._entityList = this.web.lists.getByTitle(this._params.listName)
    this._entityContentType = this._params.contentTypeId
      ? this.web.contentTypes.getById(this._params.contentTypeId)
      : null
  }

  /**
   * Configure
   *
   * @param spConfiguration - SP configuration
   */
  public configure(spConfiguration: ISPConfiguration = {}): SpEntityPortalService {
    sp.setup(spConfiguration)
    return this
  }

  /**
   * Returns a new instance of the SpEntityPortalService using the specified params
   *
   * @param params - Params
   */
  public usingParams(params: ISpEntityPortalServiceParams) {
    return new SpEntityPortalService({ ...this._params, ...params })
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
      return await query.usingCaching().get<IEntityField[]>()
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
        await this._entityList.items
          .filter(`${this._params.identityFieldName} eq '${identity}'`)
          .usingCaching()
          .get()
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
      return await this._entityList.items.getById(itemId).fieldValuesAsText.get()
    } catch (e) {
      throw e
    }
  }

  /**
   * Get entity urls
   *
   * @param {number} itemId Item id
   * @param {string} sourceUrl Source URL
   */
  protected async getEntityUrls(itemId: number, sourceUrl: string): Promise<IEntityUrls> {
    try {
      const { Id, DefaultEditFormUrl } = await this._entityList
        .select('DefaultEditFormUrl', 'Id')
        .expand('DefaultEditFormUrl')
        .usingCaching()
        .get<{ Id: string; DefaultEditFormUrl: string }>()
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
   * @param {string} identity Identity
   * @param {Object} properties Properties
   */
  public async updateEntityItem(
    identity: string,
    properties: ITypedHash<string>
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
   * @param {string} identity Identity
   * @param {string} url Url
   * @param {Object} additionalProperties Additional properties
   * @param {string} sourceUrl Source URL
   * @param {INewEntityPermissions} permissions Permissions
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
   * Set entity permissions
   *
   * @param {Item} item Item/entity
   * @param {INewEntityPermissions} permissions Permissions
   */
  private async setEntityPermissions(
    item: IItem,
    { fullControlPrincipals, readPrincipals, addEveryoneRead }: INewEntityPermissions
  ) {
    await item.breakRoleInheritance(false, true)
    if (fullControlPrincipals) {
      for (let i = 0; i < fullControlPrincipals.length; i++) {
        let principal = await this.web.ensureUser(fullControlPrincipals[i])
        await item.roleAssignments.add(principal.data.Id, 1073741829)
      }
    }
    if (readPrincipals) {
      for (let i = 0; i < readPrincipals.length; i++) {
        let principal = await this.web.ensureUser(readPrincipals[i])
        await item.roleAssignments.add(principal.data.Id, 1073741826)
      }
    }
    if (addEveryoneRead) {
      const [everyonePrincipal] = await this.web.siteUsers
        .filter(`substringof('spo-grid-all-user', LoginName)`)
        .select('Id')
        .get<Array<{ Id: number }>>()
      await item.roleAssignments.add(everyonePrincipal.Id, 1073741826)
    }
  }
}

export { ISpEntityPortalServiceParams, IEntityField, IEntity, IEntityUrls }
