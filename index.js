"use strict";
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
        this.params = params;
        this.web = new sp_1.Web(this.params.webUrl);
        this.list = this.web.lists.getByTitle(this.params.listName);
        if (this.params.contentTypeId && this.params.fieldsGroupName) {
            this.contentType = this.web.contentTypes.getById(this.params.contentTypeId);
            this.fields = this.contentType.fields.filter("Group eq '" + this.params.fieldsGroupName + "'");
        }
    }
    /**
     * Get entity item fields
     */
    SpEntityPortalService.prototype.getEntityFields = function () {
        return __awaiter(this, void 0, void 0, function () {
            var fields, e_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.fields) {
                            return [2 /*return*/, null];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.fields.get()];
                    case 2:
                        fields = _a.sent();
                        return [2 /*return*/, fields];
                    case 3:
                        e_1 = _a.sent();
                        throw e_1;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get entity item
     *
     * @param {string} groupId Group ID
     */
    SpEntityPortalService.prototype.getEntityItem = function (groupId) {
        return __awaiter(this, void 0, void 0, function () {
            var item, e_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.list.items.filter(this.params.groupIdFieldName + " eq '" + groupId + "'").get()];
                    case 1:
                        item = (_a.sent())[0];
                        return [2 /*return*/, item];
                    case 2:
                        e_2 = _a.sent();
                        throw e_2;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get entity item ID
     *
     * @param {string} groupId Group ID
     */
    SpEntityPortalService.prototype.getEntityItemId = function (groupId) {
        return __awaiter(this, void 0, void 0, function () {
            var item, e_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getEntityItem(groupId)];
                    case 1:
                        item = _a.sent();
                        return [2 /*return*/, item.Id];
                    case 2:
                        e_3 = _a.sent();
                        throw e_3;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get entity item field values
     *
     * @param {string} groupId Group ID
     */
    SpEntityPortalService.prototype.getEntityItemFieldValues = function (groupId) {
        return __awaiter(this, void 0, void 0, function () {
            var itemId, itemFieldValues, e_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getEntityItemId(groupId)];
                    case 1:
                        itemId = _a.sent();
                        return [4 /*yield*/, this.list.items.getById(itemId).fieldValuesAsText.get()];
                    case 2:
                        itemFieldValues = _a.sent();
                        return [2 /*return*/, itemFieldValues];
                    case 3:
                        e_4 = _a.sent();
                        throw e_4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
    * Get entity edit form url
    *
    * @param {string} groupId Group ID
    * @param {string} sourceUrl Source URL
    * @param {number} _itemId Item id
    */
    SpEntityPortalService.prototype.getEntityEditFormUrl = function (groupId, sourceUrl, _itemId) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, itemId, DefaultEditFormUrl, editFormUrl, e_5;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, Promise.all([
                                _itemId ? (function () { return __awaiter(_this, void 0, void 0, function () { return __generator(this, function (_a) {
                                    return [2 /*return*/, _itemId];
                                }); }); })() : this.getEntityItemId(groupId),
                                this.list.select('DefaultEditFormUrl').expand('DefaultEditFormUrl').get(),
                            ])];
                    case 1:
                        _a = _b.sent(), itemId = _a[0], DefaultEditFormUrl = _a[1].DefaultEditFormUrl;
                        editFormUrl = window.location.protocol + "//" + window.location.hostname + DefaultEditFormUrl + "?ID=" + itemId;
                        if (sourceUrl) {
                            editFormUrl += "&Source=" + encodeURIComponent(sourceUrl);
                        }
                        return [2 /*return*/, editFormUrl];
                    case 2:
                        e_5 = _b.sent();
                        throw e_5;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Update enity item
     *
     * @param {string} groupId Group ID
     * @param {Object} properties Properties
     */
    SpEntityPortalService.prototype.updateEntityItem = function (groupId, properties) {
        return __awaiter(this, void 0, void 0, function () {
            var itemId, e_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.getEntityItemId(groupId)];
                    case 1:
                        itemId = _a.sent();
                        return [4 /*yield*/, this.list.items.getById(itemId).update(properties)];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        e_6 = _a.sent();
                        throw e_6;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * New entity
     *
     * @param {PageContext} context Context
     * @param {string} sourceUrl Source URL
     */
    SpEntityPortalService.prototype.newEntity = function (context, sourceUrl) {
        if (sourceUrl === void 0) { sourceUrl = null; }
        return __awaiter(this, void 0, void 0, function () {
            var properties, data, editFormUrl, e_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        properties = { Title: context.web.title };
                        properties[this.params.groupIdFieldName] = context.legacyPageContext.groupId;
                        if (this.params.siteUrlFieldName) {
                            properties[this.params.siteUrlFieldName] = context.web.absoluteUrl;
                        }
                        return [4 /*yield*/, this.list.items.add(properties)];
                    case 1:
                        data = (_a.sent()).data;
                        return [4 /*yield*/, this.getEntityEditFormUrl(context.legacyPageContext.groupId, sourceUrl, data.Id)];
                    case 2:
                        editFormUrl = _a.sent();
                        return [2 /*return*/, { item: data, editFormUrl: editFormUrl }];
                    case 3:
                        e_7 = _a.sent();
                        throw e_7;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    return SpEntityPortalService;
}());
exports.default = SpEntityPortalService;
//# sourceMappingURL=index.js.map