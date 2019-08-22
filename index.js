"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var sp_1 = require("@pnp/sp");
var SpEntityPortalService = /** @class */ (function () {
    function SpEntityPortalService(params) {
        this.params = params;
        this._web = new sp_1.Web(this.params.webUrl);
        this._list = this._web.lists.getByTitle(this.params.listName);
        if (this.params.contentTypeId && this.params.fieldsGroupName) {
            this._contentType = this._web.contentTypes.getById(this.params.contentTypeId);
        }
    }
    /**
     * Get entity item
     *
     * @param {string} identity Identity
     * @param {string} sourceUrl Source URL used to generate URLs
     */
    SpEntityPortalService.prototype.fetchEntity = function (identity, sourceUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, item, fields, _b, urls, fieldValues;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0: return [4 /*yield*/, Promise.all([
                            this.getEntityItem(identity),
                            this.getEntityFields(),
                        ])];
                    case 1:
                        _a = _c.sent(), item = _a[0], fields = _a[1];
                        return [4 /*yield*/, Promise.all([
                                this.getEntityUrls(item.Id, sourceUrl),
                                this.getEntityItemFieldValues(item.Id),
                            ])];
                    case 2:
                        _b = _c.sent(), urls = _b[0], fieldValues = _b[1];
                        return [2 /*return*/, { item: item, fields: fields, urls: urls, fieldValues: fieldValues }];
                }
            });
        });
    };
    /**
     * Get entity fields
     */
    SpEntityPortalService.prototype.getEntityFields = function () {
        return __awaiter(this, void 0, void 0, function () {
            var e_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this._contentType) {
                            return [2 /*return*/, []];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._contentType.fields
                                .select('InternalName', 'Title', 'TypeAsString', 'SchemaXml')
                                .filter("Group eq '" + this.params.fieldsGroupName + "'")
                                .get()];
                    case 2: return [2 /*return*/, _a.sent()];
                    case 3:
                        e_1 = _a.sent();
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get entity item
     *
     * @param {string} identity Identity
     */
    SpEntityPortalService.prototype.getEntityItem = function (identity) {
        return __awaiter(this, void 0, void 0, function () {
            var e_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        if (identity.length === 38) {
                            identity = identity.substring(1, 37);
                        }
                        return [4 /*yield*/, this._list.items
                                .filter(this.params.identityFieldName + " eq '" + identity + "'")
                                .get()];
                    case 1: return [2 /*return*/, (_a.sent())[0]];
                    case 2:
                        e_2 = _a.sent();
                        throw e_2;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get entity item field values
     *
    * @param {number} itemId Item id
     */
    SpEntityPortalService.prototype.getEntityItemFieldValues = function (itemId) {
        return __awaiter(this, void 0, void 0, function () {
            var e_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._list.items
                                .getById(itemId)
                                .fieldValuesAsText
                                .get()];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2:
                        e_3 = _a.sent();
                        throw e_3;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
    * Get entity urls
    *
    * @param {number} itemId Item id
    * @param {string} sourceUrl Source URL
    */
    SpEntityPortalService.prototype.getEntityUrls = function (itemId, sourceUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, Id, DefaultEditFormUrl, editFormUrl, versionHistoryUrl, e_4;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._list
                                .select('DefaultEditFormUrl', 'Id')
                                .expand('DefaultEditFormUrl')
                                .get()];
                    case 1:
                        _a = _b.sent(), Id = _a.Id, DefaultEditFormUrl = _a.DefaultEditFormUrl;
                        editFormUrl = window.location.protocol + "//" + window.location.hostname + DefaultEditFormUrl + "?ID=" + itemId;
                        versionHistoryUrl = this.params.webUrl + "/_layouts/15/versions.aspx?list=" + Id + "&ID=" + itemId;
                        if (sourceUrl) {
                            editFormUrl += "&Source=" + encodeURIComponent(sourceUrl);
                            versionHistoryUrl = "&Source=" + encodeURIComponent(sourceUrl);
                        }
                        return [2 /*return*/, { editFormUrl: editFormUrl, versionHistoryUrl: versionHistoryUrl }];
                    case 2:
                        e_4 = _b.sent();
                        throw e_4;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Update enity item
     *
     * @param {string} identity Identity
     * @param {Object} properties Properties
     */
    SpEntityPortalService.prototype.updateEntityItem = function (identity, properties) {
        return __awaiter(this, void 0, void 0, function () {
            var item, e_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getEntityItem(identity)];
                    case 1:
                        item = _a.sent();
                        return [4 /*yield*/, this._list.items.getById(item.Id).update(properties)];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        e_5 = _a.sent();
                        throw e_5;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * New entity
     *
     * @param {string} identity Identity
     * @param {string} url Url
     * @param {Object} additionalProperties Additional properties
     * @param {string} sourceUrl Source URL
     * @param {INewEntityPermissions} permissions Permissions
     */
    SpEntityPortalService.prototype.newEntity = function (identity, url, additionalProperties, sourceUrl, permissions) {
        if (sourceUrl === void 0) { sourceUrl = null; }
        return __awaiter(this, void 0, void 0, function () {
            var _a, properties, _b, data, item, editFormUrl, e_6;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _c.trys.push([0, 5, , 6]);
                        properties = __assign((_a = {}, _a[this.params.identityFieldName] = identity, _a), additionalProperties);
                        if (this.params.urlFieldName) {
                            properties[this.params.urlFieldName] = url;
                        }
                        return [4 /*yield*/, this._list.items.add(properties)];
                    case 1:
                        _b = _c.sent(), data = _b.data, item = _b.item;
                        if (!permissions) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.setEntityPermissions(item, permissions)];
                    case 2:
                        _c.sent();
                        _c.label = 3;
                    case 3: return [4 /*yield*/, this.getEntityUrls(data.Id, sourceUrl)];
                    case 4:
                        editFormUrl = (_c.sent()).editFormUrl;
                        return [2 /*return*/, { item: data, editFormUrl: editFormUrl }];
                    case 5:
                        e_6 = _c.sent();
                        throw e_6;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Set entity permissions
     *
     * @param {Item} item Item/entity
     * @param {INewEntityPermissions} permissions Permissions
     */
    SpEntityPortalService.prototype.setEntityPermissions = function (item, _a) {
        var fullControlPrincipals = _a.fullControlPrincipals, readPrincipals = _a.readPrincipals, addEveryoneRead = _a.addEveryoneRead;
        return __awaiter(this, void 0, void 0, function () {
            var i, principal, i, principal, everyonePrincipal;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, item.breakRoleInheritance(false, true)];
                    case 1:
                        _b.sent();
                        if (!fullControlPrincipals) return [3 /*break*/, 6];
                        i = 0;
                        _b.label = 2;
                    case 2:
                        if (!(i < fullControlPrincipals.length)) return [3 /*break*/, 6];
                        return [4 /*yield*/, this._web.ensureUser(fullControlPrincipals[i])];
                    case 3:
                        principal = _b.sent();
                        return [4 /*yield*/, item.roleAssignments.add(principal.data.Id, 1073741829)];
                    case 4:
                        _b.sent();
                        _b.label = 5;
                    case 5:
                        i++;
                        return [3 /*break*/, 2];
                    case 6:
                        if (!readPrincipals) return [3 /*break*/, 11];
                        i = 0;
                        _b.label = 7;
                    case 7:
                        if (!(i < readPrincipals.length)) return [3 /*break*/, 11];
                        return [4 /*yield*/, this._web.ensureUser(readPrincipals[i])];
                    case 8:
                        principal = _b.sent();
                        return [4 /*yield*/, item.roleAssignments.add(principal.data.Id, 1073741826)];
                    case 9:
                        _b.sent();
                        _b.label = 10;
                    case 10:
                        i++;
                        return [3 /*break*/, 7];
                    case 11:
                        if (!addEveryoneRead) return [3 /*break*/, 14];
                        return [4 /*yield*/, this._web.siteUsers.filter("substringof('spo-grid-all-user', LoginName)").select('Id').get()];
                    case 12:
                        everyonePrincipal = (_b.sent())[0];
                        return [4 /*yield*/, item.roleAssignments.add(everyonePrincipal.Id, 1073741826)];
                    case 13:
                        _b.sent();
                        _b.label = 14;
                    case 14: return [2 /*return*/];
                }
            });
        });
    };
    return SpEntityPortalService;
}());
exports.SpEntityPortalService = SpEntityPortalService;
//# sourceMappingURL=index.js.map