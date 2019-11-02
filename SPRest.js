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
var SPRest;
(function (SPRest) {
    var Group = (function () {
        function Group(groupName) {
            this.groupName = groupName;
        }
        Group.prototype.getUsers = function () {
            return __awaiter(this, void 0, void 0, function () {
                var self;
                return __generator(this, function (_a) {
                    self = this;
                    return [2, new Promise(function (resolve, reject) {
                            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                                var clientContext = new SP.ClientContext.get_current();
                                var spGroup = clientContext.get_web().get_siteGroups().getByName(self.groupName);
                                var userCollection = spGroup.get_users();
                                clientContext.load(userCollection);
                                clientContext.executeQueryAsync(function () {
                                    var users = [];
                                    var userEnumerator = userCollection.getEnumerator();
                                    while (userEnumerator.moveNext()) {
                                        var oUser = userEnumerator.get_current();
                                        var user = {
                                            Id: oUser.get_id(),
                                            Title: oUser.get_title(),
                                            LoginName: oUser.get_loginName(),
                                            Email: oUser.get_email()
                                        };
                                        if (user.LoginName != "SHAREPOINT\\system") {
                                            users.push(user);
                                        }
                                    }
                                    resolve(users);
                                }, function (sender, args) {
                                    console.error(args.get_message());
                                    reject(args.get_message());
                                });
                            });
                        })];
                });
            });
        };
        Group.prototype.addUser = function (loginName) {
            return __awaiter(this, void 0, void 0, function () {
                var self;
                return __generator(this, function (_a) {
                    self = this;
                    return [2, new Promise(function (resolve, reject) {
                            var clientContext = new SP.ClientContext.get_current();
                            var web = clientContext.get_web();
                            var user = web.ensureUser(loginName);
                            var spGroup = web.get_siteGroups().getByName(self.groupName);
                            var userCollection = spGroup.get_users();
                            userCollection.addUser(user);
                            clientContext.load(user);
                            clientContext.executeQueryAsync(function () { return resolve(); }, function (sender, args) {
                                console.error(args.get_message());
                                reject(args.get_message());
                            });
                        })];
                });
            });
        };
        Group.prototype.deleteUser = function (loginName) {
            return __awaiter(this, void 0, void 0, function () {
                var self;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            self = this;
                            return [4, new Promise(function (resolve, reject) {
                                    var clientContext = new SP.ClientContext.get_current();
                                    var web = clientContext.get_web();
                                    var user = web.ensureUser(loginName);
                                    var spGroup = web.get_siteGroups().getByName(self.groupName);
                                    var userCollection = spGroup.get_users();
                                    userCollection.remove(user);
                                    clientContext.load(user);
                                    clientContext.executeQueryAsync(function () { return resolve(); }, function (sender, args) {
                                        console.error(args.get_message());
                                        reject(args.get_message());
                                    });
                                })];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        };
        return Group;
    }());
    SPRest.Group = Group;
    var User = (function () {
        function User(id) {
            if (!id)
                id = _spPageContextInfo.userId;
            this._baseUrl = location.origin + "/_api/web/siteusers/getbyid(" + id + ")";
        }
        User.prototype.get = function () {
            return __awaiter(this, void 0, void 0, function () {
                var url, response, e_1;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            url = this._baseUrl;
                            _a.label = 1;
                        case 1:
                            _a.trys.push([1, 3, , 4]);
                            return [4, $.ajax({
                                    async: true,
                                    url: url,
                                    method: "GET",
                                    headers: {
                                        "accept": "application/json;odata=nometadata"
                                    }
                                })];
                        case 2:
                            response = _a.sent();
                            return [2, response];
                        case 3:
                            e_1 = _a.sent();
                            HandleError(e_1, "User.get() failed");
                            throw Error(e_1);
                        case 4: return [2];
                    }
                });
            });
        };
        User.prototype.groups = function () {
            return __awaiter(this, void 0, void 0, function () {
                var url, response, groups, e_2;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            url = this._baseUrl + '/groups';
                            _a.label = 1;
                        case 1:
                            _a.trys.push([1, 3, , 4]);
                            return [4, $.ajax({
                                    async: true,
                                    url: url,
                                    method: "GET",
                                    headers: {
                                        "accept": "application/json;odata=nometadata"
                                    }
                                })];
                        case 2:
                            response = _a.sent();
                            groups = response.value.map(function (group) { return group.Title; });
                            return [2, groups];
                        case 3:
                            e_2 = _a.sent();
                            HandleError(e_2, "Getting user groups, roles and missions failed");
                            throw Error(e_2);
                        case 4: return [2];
                    }
                });
            });
        };
        return User;
    }());
    SPRest.User = User;
    var Delete = (function () {
        function Delete(listTitle) {
            this.listTitle = listTitle;
        }
        Delete.prototype.single = function (listItemId) {
            return this.batch([listItemId]);
        };
        Delete.prototype.folder = function (listInternalName, folderName) {
            return __awaiter(this, void 0, void 0, function () {
                var url, params;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            url = _spPageContextInfo.siteAbsoluteUrl + "/_api/web/GetFolderByServerRelativeUrl('/" + listInternalName + "/" + folderName + "')";
                            params = {
                                url: url,
                                method: "POST",
                                headers: {
                                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                    "IF-MATCH": "*",
                                    "X-HTTP-Method": "DELETE"
                                }
                            };
                            return [4, $.ajax(params)];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        };
        Delete.prototype.batch = function (listItemIds) {
            var self = this;
            return new Promise(function (resolve, reject) {
                SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                    var clientContext = new SP.ClientContext.get_current();
                    var oList = clientContext.get_web().get_lists().getByTitle(self.listTitle);
                    listItemIds.forEach(function (id) {
                        var oListItem = oList.getItemById(id);
                        oListItem.deleteObject();
                    });
                    clientContext.executeQueryAsync(resolve, onQueryFailed);
                    function onQueryFailed(sender, args) {
                        console.error(args.get_message());
                        console.error(args.get_stackTrace());
                        reject(args.get_message());
                    }
                });
            });
        };
        return Delete;
    }());
    SPRest.Delete = Delete;
    var Post = (function () {
        function Post(listTitle, listInternalName, isDocLibrary) {
            this.listInternalName = listInternalName;
            this.isDocLibrary = isDocLibrary;
            this.baseUrl = location.origin + "/_api/web/lists/getbytitle('" + listTitle + "')/items";
        }
        Object.defineProperty(Post.prototype, "params", {
            get: function () {
                return {
                    method: "POST",
                    headers: {
                        "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                        "accept": "application/json;odata=verbose",
                        "content-Type": "application/json;odata=verbose",
                    }
                };
            },
            enumerable: true,
            configurable: true
        });
        Post.prototype.prepareItem = function (listItem, fields) {
            if (fields) {
                for (var f in listItem) {
                    if (fields.indexOf(f) == -1) {
                        delete listItem[f];
                    }
                }
            }
            if (listItem.Id == null)
                delete listItem.Id;
            if (listItem.ID == null)
                delete listItem.ID;
            var append = (this.isDocLibrary) ? 'Item' : 'ListItem';
            listItem.__metadata = {
                'type': 'SP.Data.' + this.listInternalName.replace(" ", "_x0020_") + append
            };
            return JSON.stringify(listItem);
        };
        Post.prototype.createListItem = function (listItem, fields) {
            return __awaiter(this, void 0, void 0, function () {
                var url, item, params, response;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            url = this.baseUrl;
                            item = this.prepareItem(listItem, fields);
                            params = this.params;
                            params.data = item;
                            return [4, $.ajax(url, params)];
                        case 1:
                            response = _a.sent();
                            return [2, response.d.Id];
                    }
                });
            });
        };
        Post.prototype.createFolder = function (folderName) {
            return __awaiter(this, void 0, void 0, function () {
                var url;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            url = _spPageContextInfo.siteAbsoluteUrl + "/_api/web/folders/add('" + this.listInternalName + "/" + folderName + "')";
                            return [4, $.ajax(url, this.params)];
                        case 1:
                            _a.sent();
                            return [2];
                    }
                });
            });
        };
        Post.prototype.updateListItem = function (listItem, fields) {
            return __awaiter(this, void 0, void 0, function () {
                var url, item, params;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            url = this.baseUrl + ("(" + listItem.Id + ")");
                            item = this.prepareItem(listItem, fields);
                            params = this.params;
                            params.headers["IF-MATCH"] = "*";
                            params.headers["X-HTTP-Method"] = "MERGE";
                            params.data = item;
                            return [4, $.ajax(url, params)];
                        case 1:
                            _a.sent();
                            return [2, listItem.Id];
                    }
                });
            });
        };
        Post.prototype.saveListItem = function (listItem, fields) {
            if (listItem.Id) {
                return this.updateListItem(listItem, fields);
            }
            else {
                return this.createListItem(listItem, fields);
            }
        };
        return Post;
    }());
    SPRest.Post = Post;
    var Get = (function () {
        function Get(listTitle) {
            this._select = "Title,Id";
            this._expand = "";
            this._filter = "";
            this._orderby = "Title";
            this._top = 5000;
            this.baseUrl = location.origin + "/_api/web/lists/getbytitle('" + listTitle + "')";
        }
        Get.prototype.select = function (select) {
            this._select = select;
            return this;
        };
        Get.prototype.expand = function () {
            var fields = this._select.split(',');
            var expand = fields
                .filter(function (field) { return field.includes('/'); })
                .map(function (field) { return field.substring(0, field.indexOf('/')); });
            this._expand = _.uniq(expand).join(',');
        };
        Get.prototype.filter = function (filter) {
            if (Array.isArray(filter)) {
                filter = filter.map(function (Id) {
                    return "Id eq " + Id;
                }).join(" or ");
            }
            this._filter = filter;
            return this;
        };
        Get.prototype.orderBy = function (orderby) {
            this._orderby = orderby;
            return this;
        };
        Get.prototype.top = function (limit) {
            this._top = limit;
            return this;
        };
        Get.prototype.url = function () {
            var url = this.baseUrl + "/items?$select=" + this._select;
            this.expand();
            if (this._expand)
                url += "&$expand=" + this._expand;
            if (this._filter)
                url += "&$filter=" + this._filter;
            url += "&$orderby=" + this._orderby + "&$top=" + this._top;
            return url;
        };
        Get.prototype.items = function () {
            return __awaiter(this, void 0, void 0, function () {
                var self_1, response, e_3;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            self_1 = this;
                            return [4, $.ajax({
                                    async: true,
                                    url: self_1.url(),
                                    method: "GET",
                                    headers: {
                                        "accept": "application/json;odata=nometadata"
                                    }
                                })];
                        case 1:
                            response = _a.sent();
                            return [2, response.value];
                        case 2:
                            e_3 = _a.sent();
                            HandleError(e_3, "Get items failed for query " + this.url());
                            throw Error(e_3);
                        case 3: return [2];
                    }
                });
            });
        };
        Get.prototype.titles = function () {
            return __awaiter(this, void 0, void 0, function () {
                var resultItems;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0: return [4, this.items()];
                        case 1:
                            resultItems = _a.sent();
                            return [2, resultItems.map(function (item) { return item.Title; })];
                    }
                });
            });
        };
        Get.prototype.find = function (id) {
            return __awaiter(this, void 0, void 0, function () {
                return __generator(this, function (_a) {
                    this._filter = "Id eq " + id;
                    return [2, this.single()];
                });
            });
        };
        Get.prototype.single = function () {
            return __awaiter(this, void 0, void 0, function () {
                var self_2, response, e_4;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            self_2 = this;
                            return [4, $.ajax({
                                    async: true,
                                    cache: false,
                                    url: self_2.url(),
                                    method: "GET",
                                    headers: {
                                        "accept": "application/json;odata=nometadata"
                                    }
                                })];
                        case 1:
                            response = _a.sent();
                            if (response.value && response.value.length == 1) {
                                return [2, response.value[0]];
                            }
                            else {
                                console.warn(this.url());
                                throw Error("Get single list item failed.");
                            }
                            return [3, 3];
                        case 2:
                            e_4 = _a.sent();
                            HandleError(e_4, "Get single/find failed for query " + this.url());
                            throw Error(e_4);
                        case 3: return [2];
                    }
                });
            });
        };
        return Get;
    }());
    SPRest.Get = Get;
    var CAML = (function () {
        function CAML(listTitle) {
            this.viewfields = '';
            this.where = '';
            this.url = location.origin + "/_api/web/lists/getbytitle('" + listTitle + "')/GetItems";
        }
        CAML.prototype.params = function (query) {
            return {
                method: "POST",
                headers: {
                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                    "accept": "application/json;odata=nometadata",
                    "content-Type": "application/json;odata=nometadata",
                },
                data: JSON.stringify({
                    query: {
                        ViewXml: query
                    }
                })
            };
        };
        CAML.prototype.select = function (fields) {
            var viewfields = fields.split(",").map(function (field) { return "<FieldRef Name=\"" + field + "\"></FieldRef>"; }).join();
            this.viewfields = "<ViewFields>" + viewfields + "</ViewFields>";
            return this;
        };
        CAML.prototype.filter = function (where) {
            where = where.replace(/\s{2,}/gm, ' ');
            this.where = "<Where>" + where + "</Where>";
            return this;
        };
        CAML.prototype.items = function () {
            return __awaiter(this, void 0, void 0, function () {
                var query, caml, viewXml, response, e_5;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            _a.trys.push([0, 2, , 3]);
                            query = "<Query>" + this.where + "</Query>";
                            caml = "<View>" + this.viewfields + query + "</View>";
                            viewXml = JSON.stringify({ ViewXml: caml });
                            console.info(viewXml);
                            return [4, $.ajax(this.url, this.params(viewXml))];
                        case 1:
                            response = _a.sent();
                            return [2, response.value];
                        case 2:
                            e_5 = _a.sent();
                            HandleError(e_5, "CAML get failed");
                            throw Error(e_5);
                        case 3: return [2];
                    }
                });
            });
        };
        return CAML;
    }());
    SPRest.CAML = CAML;
    function HandleError(response, prependMessage) {
        try {
            var readableError = "";
            if (response.responseText) {
                var o = JSON.parse(response.responseText);
                if (o.error && o.error.message) {
                    readableError = o.error.message.value;
                }
                else if (o["odata.error"] && o["odata.error"].message) {
                    readableError = o["odata.error"].message.value;
                }
                if (readableError) {
                    console.error(readableError);
                    return readableError;
                }
            }
            else if (response.message && response.stack) {
                console.error(response.stack);
                return response.message;
            }
            console.error(response);
            return 'Error: ' + prependMessage;
        }
        catch (e) {
            console.error(prependMessage, e);
            return 'Error: ' + prependMessage;
        }
    }
})(SPRest || (SPRest = {}));
//# sourceMappingURL=SPRest.js.map