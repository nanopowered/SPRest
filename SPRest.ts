namespace SPRest {

    //#region Base Interfaces & classes
    interface IItem {
        Id: number;
        Title: string;
    }

    interface IUser {
        Id: number, LoginName: string, Title: string, Email: string
    }

    interface IPerson {
        Id: number, Name: string, Title: string, EMail: string
    }
    /**
     * Created and Modified can have type string or Date
     * if the type is string, the string is in an iso format 
     * which can be cast to Date this way: new Date(string_ISO_format)
     */
    interface IlistItem extends IItem {
        // If the type is string, it should be a date in ISO format
        Created: string | Date;
        // If the type is string, it should be a date in ISO format
        Modified: string | Date;
        Author: IPerson;
        Editor: IPerson;
    }
    //#endregion

    //#region Group
    /**
     * Get, add and delete users from group
     */
    export class Group {
        // use of client object model for groups instead of REST because with REST
        // there are sometimes credentials popup for non admin group, even if they have the appropriate authorizations
        private groupName: string;

        constructor(groupName: string) {
            this.groupName = groupName;
        }

        async getUsers(): Promise<any> {
            let self = this;
            return new Promise((resolve, reject) => {
                SP.SOD.executeFunc('sp.js', 'SP.ClientContext', () => {
                    let clientContext = new SP.ClientContext.get_current();
                    let spGroup = clientContext.get_web().get_siteGroups().getByName(self.groupName)
                    let userCollection = spGroup.get_users();
                    clientContext.load(userCollection);
                    clientContext.executeQueryAsync(() => {
                        // query succeeded
                        let users: Array<IUser> = [];
                        var userEnumerator = userCollection.getEnumerator();
                        while (userEnumerator.moveNext()) {
                            var oUser = userEnumerator.get_current();
                            let user: IUser = {
                                Id: oUser.get_id(),
                                Title: oUser.get_title(),
                                LoginName: oUser.get_loginName(),
                                Email: oUser.get_email()
                            };
                            // we are not interested in system account which is owner in many groups
                            if (user.LoginName != "SHAREPOINT\\system") {
                                users.push(user);
                            }
                        }
                        resolve(users);
                    }, (sender: any, args: any) => {
                        console.error(args.get_message());
                        reject(args.get_message());
                    });
                });
            });
        }

        async addUser(loginName: string): Promise<any> {
            let self = this;
            return new Promise((resolve, reject) => {
                let clientContext = new SP.ClientContext.get_current();
                let web = clientContext.get_web();
                let user = web.ensureUser(loginName);
                let spGroup = web.get_siteGroups().getByName(self.groupName);
                let userCollection = spGroup.get_users();
                userCollection.addUser(user);
                clientContext.load(user);
                clientContext.executeQueryAsync(
                    () => resolve(),
                    (sender: any, args: any) => {
                        console.error(args.get_message());
                        reject(args.get_message());
                    }
                );
            });
        }

        async deleteUser(loginName: string): Promise<void> {
            let self = this;
            await new Promise((resolve, reject) => {
                let clientContext = new SP.ClientContext.get_current();
                let web = clientContext.get_web();
                let user = web.ensureUser(loginName);
                let spGroup = web.get_siteGroups().getByName(self.groupName);
                let userCollection = spGroup.get_users();
                userCollection.remove(user);
                clientContext.load(user);
                clientContext.executeQueryAsync(
                    () => resolve(),
                    (sender: any, args: any) => {
                        console.error(args.get_message());
                        reject(args.get_message());
                    }
                );
            });
        }

    }
    //#endregion

    //#region User
    /**
     * @param id user id, if undefined, we use the id of the current user
     */
    export class User {
        private _baseUrl: string;

        constructor(id?: number) {
            // if the id is not provided, we use the current user
            if (!id) id = _spPageContextInfo.userId;
            this._baseUrl = `${location.origin}/_api/web/siteusers/getbyid(${id})`;
        }

        /**
         * Get informations on user
         */
        async get() {
            let url = this._baseUrl;
            try {
                let response = await $.ajax({
                    async: true,
                    url: url,
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=nometadata"
                    }
                });
                return response;
            } catch (e) {
                HandleError(e, `User.get() failed`);
                throw Error(e);
            }
        }

        /**
         * Get user's groups
         */
        async groups() {
            let url = this._baseUrl + '/groups';
            try {
                let response = await $.ajax({
                    async: true,
                    url: url,
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=nometadata"
                    }
                });
                let groups = response.value.map((group: any) => group.Title);
                return groups;
            } catch (e) {
                HandleError(e, `Getting user groups, roles and missions failed`);
                throw Error(e);
            }
        }
    }
    //#endregion

    //#region Delete
    /**
     * @param {string} listTitle - the list display name / title
     */
    export class Delete {
        private listTitle: string;

        constructor(listTitle: string) {
            this.listTitle = listTitle;
        }

        /**
         * @description delete a list item
         * @param listItemId - id of the list item
         */
        single(listItemId: number): Promise<void> {
            return this.batch([listItemId]);
        }

        /**
         * @param {string} listInternalName - the internal (or static) name of the list
         * @param {string} folderName - the name of the folder to delete
         */
        async folder(listInternalName: string, folderName: string): Promise<void> {
            let url = `${_spPageContextInfo.siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('/${listInternalName}/${folderName}')`;
            let params: any = {
                url: url,
                method: "POST",
                headers: {
                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                    "IF-MATCH": "*",
                    "X-HTTP-Method": "DELETE"
                }
            };
            await $.ajax(params);
        }

        /**
         * @description delete listitems in batch
         * @param listItemIds - an array of Ids
         */
        batch(listItemIds: Array<number>): Promise<void> {
            let self = this;
            return new Promise(function (resolve, reject) {
                SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                    var clientContext = new SP.ClientContext.get_current();
                    var oList = clientContext.get_web().get_lists().getByTitle(self.listTitle);
                    listItemIds.forEach(id => {
                        var oListItem = oList.getItemById(id);
                        oListItem.deleteObject();
                    })
                    clientContext.executeQueryAsync(resolve, onQueryFailed);

                    function onQueryFailed(sender: any, args: any) {
                        console.error(args.get_message());
                        console.error(args.get_stackTrace());
                        reject(args.get_message());
                    }
                });
            });
        }
    }
    //#endregion

    //#region Post
    /** 
     * @param {string} listTitle - the list display name / title
     * @param {string} listInternalName - the list internal (or static) name
     * @param {boolean|undefined} isDocLibrary - is the list a document library (default is false)
     * @example
     * new Post('Toys and games','ToysAndGames')
     * .createListItem({Title:'skateboard', Price: 18})
     * .then(id => console.log(id))
     */
    export class Post {
        private baseUrl: string;
        private listInternalName: string;
        private isDocLibrary?: boolean;
        private get params(): any {
            return {
                method: "POST",
                headers: {
                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                    "accept": "application/json;odata=verbose",
                    "content-Type": "application/json;odata=verbose",
                }
            }
        }

        constructor(listTitle: string, listInternalName: string, isDocLibrary?: boolean) {
            this.listInternalName = listInternalName;
            this.isDocLibrary = isDocLibrary;
            this.baseUrl = `${location.origin}/_api/web/lists/getbytitle('${listTitle}')/items`;
        }

        // add the metadata
        // if fields are provided, keep only the listItem properties listed in fields 
        private prepareItem(listItem: any, fields?: Array<string>) {
            if (fields) {
                for (let f in listItem) {
                    if (fields.indexOf(f) == -1) {
                        delete listItem[f];
                    }
                }
            }
            if (listItem.Id == null) delete listItem.Id;
            if (listItem.ID == null) delete listItem.ID;
            let append = (this.isDocLibrary) ? 'Item' : 'ListItem';
            listItem.__metadata = {
                'type': 'SP.Data.' + this.listInternalName.replace(" ", "_x0020_") + append
            }
            return JSON.stringify(listItem);
        }

        /** 
         * @param {object} listItem - a regular object, whose properties are fields' names
         * @param {array} fields - item's fields to insert, optionnal
         * @returns Promise object which represents the result, the id of the listitem
         */
        async createListItem(listItem: any, fields?: Array<string>): Promise<number> {
            let url = this.baseUrl;
            let item = this.prepareItem(listItem, fields);
            let params = this.params;
            params.data = item;
            let response = await $.ajax(url, params);
            return response.d.Id as number;
        }

        /**
         * @param {object} folderName - athe name of the folder to create
         * @returns {Promise} Promise object represents the result
         */
        async createFolder(folderName: string): Promise<void> {
            let url = `${_spPageContextInfo.siteAbsoluteUrl}/_api/web/folders/add('${this.listInternalName}/${folderName}')`;
            await $.ajax(url, this.params);
        }

        /**
         * @param {object} listItem - a regular object, whose properties are fields' names
         * @param {array}  fields - item's fields to update, optionnal
         * @returns {Promise} Promise object represents the result, the id of the listitem
         */
        async updateListItem(listItem: any, fields?: Array<string>): Promise<number> {
            let url = this.baseUrl + `(${listItem.Id})`;
            let item = this.prepareItem(listItem, fields);
            let params = this.params;
            params.headers["IF-MATCH"] = "*";
            params.headers["X-HTTP-Method"] = "MERGE";
            params.data = item;
            await $.ajax(url, params);
            return listItem.Id;
        }

        /**
         * Save a list item
         * If an Id is provided, it will update the list item,
         * otherwise it will create it
         * @param {object} listItem - a regular object, whose properties are fields' names
         * @param {array}  fields - item's fields to save
         * @returns {Promise} Promise object represents the result, the id of the listitem
         */
        saveListItem(listItem: any, fields?: Array<string>): Promise<number> {
            if (listItem.Id) {
                return this.updateListItem(listItem, fields);
            } else {
                return this.createListItem(listItem, fields);
            }
        }
    }
    //#endregion

    //#region Get
    /**
     * @param {string} listTitle - the list display name / title
     * @remarks the chain of methods must end with single(), find(id), titles() or items() to trigger the request
     * @remarks expand is automatically generated using the select fields
     * @example - get a user info
     * new Get('User Information List').select("Id,Title,Name,EMail").filter(`Name eq '${encodeURIComponent(loginName)}'`).single().then(console.log})
     * @example - get a country by id
     * new Get('Countries').find(12).then(console.log)
     * @example - filter countries by the id of a mission (lookup column)
     * new Get('Countries').select("Title,Id,Mission/Id").filter("Mission/Id eq 1").items().then(console.log);
     */
    export class Get {
        constructor(listTitle: string) {
            this.baseUrl = `${location.origin}/_api/web/lists/getbytitle('${listTitle}')`
        }

        private baseUrl: string;
        private _select: string = "Title,Id";
        private _expand: string = "";
        private _filter: string = "";
        private _orderby: string = "Title";
        private _top: number = 5000;

        /** 
         * @param select - comma separated list of fieds - default value: "Title,Id"
         * @example "Title,Id,Contract/Id,Contract/Title"
         */
        select(select: string): Get {
            this._select = select;
            return this;
        }

        /** 
         * Parse _select and generate the expand fields if needed
         */
        private expand() {
            let fields: string[] = this._select.split(',');
            let expand: string[] = fields
                .filter((field: string) => field.includes('/'))
                .map((field: string) => field.substring(0, field.indexOf('/')));
            this._expand = _.uniq(expand).join(',');
        }

        /** 
         * @param filter - filter expression or an array of ids
         * @example filter expression: "(Unassigned eq 0) and (Mission/Id eq 1)"
         * @example ids array: [1,17,128,56]
         */
        filter(filter: string | Array<number>): Get {
            if (Array.isArray(filter)) {
                filter = filter.map(function (Id) {
                    return "Id eq " + Id;
                }).join(" or ");
            }
            this._filter = filter;
            return this;
        }

        /** 
         * @param orderby - comma separated list of fieds - default value: "Title"
         */
        orderBy(orderby: string): Get {
            this._orderby = orderby;
            return this;
        }

        /** 
         * @param limit - max number of rows - default is 5000
         */
        top(limit: number): Get {
            this._top = limit;
            return this;
        }

        private url(): string {
            let url = `${this.baseUrl}/items?$select=${this._select}`;
            this.expand();
            if (this._expand) url += "&$expand=" + this._expand;
            if (this._filter) url += "&$filter=" + this._filter;
            url += `&$orderby=${this._orderby}&$top=${this._top}`;
            return url;
        }

        /** 
         * @returns Promise object which represents the result, an array of list items
         */
        async items(): Promise<any> {
            try {
                let self = this;
                let response = await $.ajax({
                    async: true,
                    url: self.url(),
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=nometadata"
                    }
                });
                return response.value;
            } catch (e) {
                HandleError(e, "Get items failed for query " + this.url());
                throw Error(e);
            }
        }

        /** 
         * @returns Promise object which represents the result, a flat array of titles
         */
        async titles(): Promise<string[]> {
            let resultItems = await this.items();
            return resultItems.map((item: { Title: string }) => item.Title);
        }

        /** 
         * @returns Promise object which represents the result, a list item
         */
        async find(id: number): Promise<any> {
            this._filter = `Id eq ${id}`;
            return this.single();
        }

        /** 
         * @returns Promise object which represents the result, a list item
         */
        async single(): Promise<any> {
            try {
                let self = this;
                let response = await $.ajax({
                    async: true,
                    cache: false,
                    url: self.url(),
                    method: "GET",
                    headers: {
                        "accept": "application/json;odata=nometadata"
                    }
                });
                if (response.value && response.value.length == 1) {
                    return response.value[0];
                } else {
                    console.warn(this.url());
                    throw Error("Get single list item failed.");
                }
            } catch (e) {
                HandleError(e, "Get single/find failed for query " + this.url());
                throw Error(e);
            }
        }
    }
    //#endregion

    //#region CAML
    /** 
     * @param {string} listTitle - the list display name / title
     * @example
     * new CAML('Questions')
     * .select("Title,DateOfExpiry")
     * .filter(`<Geq><FieldRef Name='DateOfExpiry' /><Value Type='DateTime'><Today /></Value></Geq>`)
     * .items().then(items => console.log(items))
     */
    export class CAML {
        private url: string;
        private viewfields: string = '';
        private where: string = '';

        private params(query: string): any {
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
            }
        }

        constructor(listTitle: string) {
            this.url = `${location.origin}/_api/web/lists/getbytitle('${listTitle}')/GetItems`;
        }

        /** 
         * @param select - comma separated list of fieds - default value: "Title,Id"
         * @example "Title,DateOfExpiry"
         * @remarks Id, Created, Modified, FileSystemObjectType are always included
         */
        select(fields: string): CAML {
            let viewfields = fields.split(",").map(field => `<FieldRef Name="${field}"></FieldRef>`).join();
            this.viewfields = `<ViewFields>${viewfields}</ViewFields>`;
            return this;
        }

        /** 
         * @param where - a where caml expression
         * @example 
         * `<And>
         *     <And>
         *         <Eq><FieldRef Name='Guide' LookupId='TRUE' /><Value Type='Lookup'>${id}</Value></Eq>
         *         <Geq><FieldRef Name='DateOfExpiry' /><Value Type='DateTime'><Today /></Value></Geq>
         *     </And>
         *     <Eq><FieldRef Name='ExpiryReminder'/><Value Type='Boolean'>1</Value></Eq>
         * </And>`
         */
        filter(where: string): CAML {
            // reduce whitespaces
            where = where.replace(/\s{2,}/gm, ' ');
            this.where = `<Where>${where}</Where>`;
            return this;
        }

        /** 
         * @returns Promise object which represents the result, an array of list items
         */
        async items(): Promise<any> {
            try {
                //build caml
                let query = `<Query>${this.where}</Query>`;
                let caml = `<View>${this.viewfields}${query}</View>`; //
                let viewXml = JSON.stringify({ ViewXml: caml });

                console.info(viewXml)
                // ajax call
                let response = await $.ajax(this.url, this.params(viewXml));
                return response.value;
            } catch (e) {
                HandleError(e, "CAML get failed");
                throw Error(e);
            }
        }
    }
    //#endregion

    //#region HandleError
    /**
     * 
     * @param response Handle ajax errors regardless of their formats
     * @param prependMessage text at the beginning
     * @returns a message
     */
    function HandleError(response: any, prependMessage: string): string {
        try {
            let readableError = ""
            if (response.responseText) {
                let o = JSON.parse(response.responseText);
                if (o.error && o.error.message) {
                    readableError = o.error.message.value;
                } else if (o["odata.error"] && o["odata.error"].message) {
                    readableError = o["odata.error"].message.value;
                }
                if (readableError) {
                    console.error(readableError);
                    return readableError;
                }
            } else if (response.message && response.stack) {
                console.error(response.stack);
                return response.message;
            }
            console.error(response);
            return 'Error: ' + prependMessage;
        } catch (e) {
            console.error(prependMessage, e);
            return 'Error: ' + prependMessage;
        }
    }
    //#endregion
}
