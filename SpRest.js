var SPRest = /** @class */ (function() {
    function SPRest(rootWeb) {
        var _this = this;
        this.Utils = (function() {
            /**
             * Reference or motivation link : https://github.com/omkarkhair/sp-jello/blob/master/lib/jello.js
             */
            var reqUrl = _this.rootUrl + "/_api/web/lists/getbytitle('AWSResponseList')/items";
            //var fetch = require("node-fetch");
            // TODO- Need to transform this json to closure function to allow only get operation.
            //TODO  - Need to add support for Field Meta in utils
            var spDefaultMeta = {
                List: {
                    BaseTemplate: 100,
                    // "Title": "List Created with Fetch UTIL1112",
                    __metadata: { type: "SP.List" }
                },
                Document: {
                    __metadata: { type: "SP.List" },
                    AllowContentTypes: true,
                    BaseTemplate: 101,
                    ContentTypesEnabled: true
                        // 'Description': 'My doc. lib. description',
                        // 'Title': 'Test'
                },
                ListItem: {
                    __metadata: { type: "SP.Data.AWSResponseListListItem" }
                    // "Title": "Test"
                }
            };
            console.log(_this.rootUrl);
            var isOdataVerbose = true;
            var _rootUrl = _this.rootUrl;
            //Here "Accept" and "Content-Type" header must require for Sharepoint Onlin REST API
            var _payloadOptions = {
                method: "GET",
                body: undefined,
                headers: {
                    credentials: "include",
                    Accept: "application/json; odata=verbose",
                    "Content-Type": "application/json; odata=verbose"
                }
            };
            // TODO- To support for param base odata
            //No to fetch Metadata, response only requested data.
            // Reference link https://www.microsoft.com/en-us/microsoft-365/blog/2014/08/13/json-light-support-rest-sharepoint-api-released/
            //Option 1: verbose “accept: application/json; odata=verbose”
            //Option 2: minimalmetadata “accept: application/json; odata=minimalmetadata”
            //Option 3: nometadata “accept: application/json; odata=nometadata”
            //Option 4: Don’t provide it “accept: application/json” This defaults to minimalmetadata option
            if (!isOdataVerbose) {
                _payloadOptions.headers.Accept = "application/json; odata=nometadata";
            }
            // Reference from :https://www.tjvantoll.com/2015/09/13/fetch-and-errors/
            //Reference rom : https://sharepoint.stackexchange.com/questions/105380/adding-new-list-item-using-rest
            // Get List Item Type metadata
            var getItemTypeForListName = function(name) {
                return ("SP.Data." +
                    name.charAt(0).toUpperCase() +
                    name
                    .split(" ")
                    .join("")
                    .slice(1) +
                    "ListItem");
            };
            var Get = function(url) {
                var _localPayload = _payloadOptions;
                _localPayload.method = "GET";
                //Internally if body is set for GET request then need to remove it by setting undefined
                // otherwise it return with error : Failed to execute 'fetch' on 'Window': Request with GET/HEAD method cannot have body
                _localPayload.body = undefined;
                delete _localPayload.headers["IF-MATCH"];
                delete _localPayload.headers["X-HTTP-Method"];
                console.log(_localPayload);
                return fetch(url, _localPayload).then(function(r) { return r.json(); });
            };
            var getRequestDigest = function(url) {
                if (url === void 0) { url = ""; }
                if (url) {
                    url = _this.rootUrl;
                }
                url += "/_api/contextinfo";
                var _localPayload = _payloadOptions;
                _localPayload.method = "POST";
                return fetch(url, _payloadOptions).then(function(r) { return r.json(); });
            };
            var postWithRequestDigestExtension = function(url, _a) {
                var _b = _a.headers,
                    headers = _b === void 0 ? {} : _b,
                    payload = _a.payload;
                return getRequestDigest().then(function(token) {
                    // payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
                    headers["X-RequestDigest"] =
                        token.d.GetContextWebInformation.FormDigestValue;
                    return PostExtension(url, { headers: headers, payload: payload });
                });
            };
            var postWithRequestDigest = function(url, payload) {
                return getRequestDigest().then(function(token) {
                    payload.requestDigest =
                        token.d.GetContextWebInformation.FormDigestValue;
                    return Post(url, payload);
                });
            };
            //Need to refactor below method and need to merge in single method postWithRequestDigestExtension
            //Based on Action Type in headers using switch case add or remove extra headers
            // Try to add header name same as standard name so it can be replaced with for loop e.g. _extraHeaders
            var updateWithRequestDigest = function(url, payload) {
                return getRequestDigest(payload.rootUrl + "/_api/contextinfo").then(function(token) {
                    payload.requestDigest =
                        token.d.GetContextWebInformation.FormDigestValue;
                    payload._extraHeaders = {
                        "IF-MATCH": "*",
                        "X-HTTP-Method": "MERGE"
                    };
                    //For Update operation or Merge Operation no response will return only status will return for http request
                    payload.isNoJsonResponse = true;
                    return Post(url, payload);
                });
            };
            var deleteWithRequestDigest = function(url, payload) {
                return getRequestDigest(payload.rootUrl + "/_api/contextinfo").then(function(token) {
                    payload.requestDigest =
                        token.d.GetContextWebInformation.FormDigestValue;
                    payload._extraHeaders = {
                        "IF-MATCH": "*",
                        "X-HTTP-Method": "DELETE"
                    };
                    //For Update operation or Merge Operation no response will return only status will return for http request
                    payload.isNoJsonResponse = true;
                    return Post(url, payload);
                });
            };
            var Post = function(url, payload) {
                var _localPayload = _payloadOptions;
                // TODO For Safety this method can be wrapped with request Digest so always get token.
                // But need to ensure it request only when request digest is expired.
                _localPayload.method = "POST";
                _localPayload.body = payload.data;
                var _metaInfo = payload.metaInfo;
                //Pre validation Check Before update body or meta detail
                if (_metaInfo && spDefaultMeta[_metaInfo.type]) {
                    //Update Title and Description while creating new List/Column/Fields
                    var type = _metaInfo.type,
                        title = _metaInfo.title,
                        listName = _metaInfo.listName;
                    var _body = spDefaultMeta[type];
                    if (title) {
                        //Update title if it is present in metaInfo
                        _body.Title = title;
                    }
                    if (type === "ListItem") {
                        //TODO Extra Efforts
                        _body.__metadata.type = getItemTypeForListName(listName);
                    }
                    //If Extra fields are present then store to payload
                    //Before Stringify ,store extra fields/columns in body
                    if (payload.extraFields) {
                        for (var field in payload.extraFields) {
                            //Add all extra columns in body
                            // TODO- Need to ensure that column value is stringify
                            _body[field] = payload.extraFields[field];
                        }
                        // Merge two object
                        // _localPayload.body = Object.assign({}, _localPayload.body, payload.extraFields)
                    }
                    //Pass body data as stringyfy;
                    _localPayload.body = JSON.stringify(_body);
                    _localPayload.body.Title = title;
                }
                //If Extra header is present in payload for update or other operation, Append to existing header
                if (payload._extraHeaders) {
                    for (var _header in payload._extraHeaders) {
                        _localPayload.headers[_header] = payload._extraHeaders[_header];
                    }
                } else {
                    //IF Not present that means it is POST Request, reset Extraheader for this request
                    _localPayload.headers["IF-MATCH"] = undefined;
                    _localPayload.headers["X-HTTP-Method"] = undefined;
                }
                _localPayload.headers["X-RequestDigest"] = payload.requestDigest;
                console.log(_localPayload);
                //TODO- Naming convention can be updated.
                if (payload.isNoJsonResponse) {
                    return fetch(url, _localPayload).then(function(r) { return r; });
                } else {
                    return fetch(url, _localPayload).then(function(r) { return r.json(); });
                }
            };
            var PostExtension = function(url, _a) {
                var headers = _a.headers,
                    payload = _a.payload;
                var _localPayload = {};
                _localPayload = _payloadOptions;
                // TODO For Safety this method can be wrapped with request Digest so always get token.
                // But need to ensure it request only when request digest is expired.
                _localPayload.method = "POST";
                //If Extra header is present in payload for update or other operation, Append to existing header
                if (headers) {
                    for (var _header in headers) {
                        _localPayload.headers[_header] = headers[_header];
                    }
                }
                //Below code not required to reset functionality
                //    else {
                //     //IF Not present that means it is POST Request, reset Extraheader for this request
                //     __copyPayloadOptions.headers["IF-MATCH"] = undefined;
                //     __copyPayloadOptions.headers["X-HTTP-Method"] = undefined;
                //   }
                //TODO -Assume payload is in object so applied stringyfy other wise need check before apply
                _localPayload.body = JSON.stringify(payload.data);
                // console.log(_payloadOptions);
                //TODO- Naming convention can be updated.
                // if (payload.isNoJsonResponse) {
                //Instead of storing  extra info in payload use headers
                if (payload.isNoJsonResponse) {
                    return fetch(url, _localPayload).then(function(r) { return r; });
                } else {
                    return fetch(url, _localPayload).then(function(r) { return r.json(); });
                }
            };
            var generateRestRequest = function(_a) {
                var _b = _a.listName,
                    listName = _b === void 0 ? "" : _b,
                    _c = _a.Id,
                    Id = _c === void 0 ? "" : _c,
                    type = _a.type;
                var prepareRequest = _this.rootUrl + "/_api/web/lists";
                switch (type) {
                    case "ListItem":
                        prepareRequest += "/getbytitle('" + listName + "')/items";
                        if (Id) {
                            prepareRequest += "(" + Id + ")";
                        }
                        break;
                    case "List":
                        if (listName) {
                            prepareRequest += "/getbytitle('" + listName + "')";
                        }
                        //TODO
                        break;
                }
                //OLDER Code for ListItem
                //   prepareRequest = `/getbytitle('${listName}')/items`;
                //   if (Id) {
                //     prepareRequest += `(${Id})`;
                //   }
                return prepareRequest;
            };
            //Provide support for odata query so, for specific provided expression append it with base requet
            var ListItem = {
                Add: function(_a) {},
                Update: function(_a) {},
                Delete: function(_a) {},
                GetItemById: function(_a) {},
                GetAllItem: function(_a) {}
            };
            var List = {
                Add: function(_a) {},
                Update: function(_a) {
                    console.log("Implementation is pending");
                },
                Delete: function(_a) {
                    console.log("Implementation is pending");
                },
                GetItemById: function(_a) {
                    console.log("Implementation is pending");
                },
                GetAll: function(_a) {}
            };
            List.GetAll = function() {
                return Get(generateRestRequest({ listName: "", type: "List" }));
            };
            List.Add = function(_a) {
                var listName = _a.listName,
                    data = _a.data;
                var _reqUrl = generateRestRequest({ listName: listName, type: "List" });
                var _prePayload = preparePayloadData({
                    action: "ADD",
                    type: "List",
                    listName: listName,
                    data: data
                });
                return postWithRequestDigestExtension(_reqUrl, _prePayload);
            };
            ListItem.GetAllItem = function(_a) {
                var listName = _a.listName;
                return Get(generateRestRequest({ listName: listName, type: "ListItem" }));
            };
            ListItem.GetItemById = function(_a) {
                var listName = _a.listName,
                    Id = _a.Id;
                return Get(generateRestRequest({ listName: listName, Id: Id, type: "ListItem" }));
            };
            // Add/Update/Delete require same types of metaData as well as payload only difference is Headers
            ListItem.Add = function(_a) {
                var listName = _a.listName,
                    data = _a.data;
                var _reqUrl = generateRestRequest({ listName: listName, type: "ListItem" });
                var _prePayload = preparePayloadData({
                    action: "ADD",
                    type: "ListItem",
                    listName: listName,
                    data: data
                });
                return postWithRequestDigestExtension(_reqUrl, _prePayload);
            };
            ListItem.Update = function(_a) {
                var listName = _a.listName,
                    Id = _a.Id,
                    data = _a.data;
                var _reqUrl = generateRestRequest({ listName: listName, Id: Id, type: "ListItem" });
                var _prePayload = preparePayloadData({
                    action: "UPDATE",
                    type: "ListItem",
                    listName: listName,
                    data: data
                });
                return postWithRequestDigestExtension(_reqUrl, _prePayload);
            };
            ListItem.Delete = function(_a) {
                var listName = _a.listName,
                    Id = _a.Id,
                    data = _a.data;
                var _reqUrl = generateRestRequest({ listName: listName, Id: Id, type: "ListItem" });
                var _prePayload = preparePayloadData({
                    action: "DELETE",
                    type: "ListItem",
                    listName: listName,
                    data: data
                });
                return postWithRequestDigestExtension(_reqUrl, _prePayload);
            };
            var preparePayloadData = function(_a) {
                var listName = _a.listName,
                    data = _a.data,
                    action = _a.action,
                    type = _a.type;
                var payload = {
                    data: { __metadata: spDefaultMeta[type] }
                };
                //payload.data = {};
                //   let _metaInfo: any = spDefaultMeta["ListItem"];
                //   _metaInfo.__metadata.type = getItemTypeForListName(listName);
                // payload.data.__metadata=spDefaultMeta["ListItem"];
                switch (type) {
                    case "ListItem":
                        //Update ListItem Content type
                        payload.data.__metadata.type = getItemTypeForListName(listName);
                        break;
                }
                //Need to check is data is object then iterate otherwise not
                if (data) {
                    for (var _key in data) {
                        payload.data[_key] = data[_key];
                    }
                }
                //TODO Verify about extra header
                // //Pass body data as stringyfy;
                // // _payloadOptions.body = JSON.stringify(_body);
                // // _payloadOptions.body.Title = title;
                //    return postWithRequestDigest(_reqUrl);
                var _headers = {};
                switch (action) {
                    case "ADD":
                        _headers = {
                            "IF-MATCH": undefined,
                            "X-HTTP-Method": undefined
                        };
                        break;
                    case "UPDATE":
                        _headers = {
                            "IF-MATCH": "*",
                            "X-HTTP-Method": "MERGE"
                        };
                        //For Update operation or Merge Operation no response will return only status will return for http request
                        payload.isNoJsonResponse = true;
                        break;
                    case "DELETE":
                        _headers = {
                            "IF-MATCH": "*",
                            "X-HTTP-Method": "DELETE"
                        };
                        //For Update operation or Merge Operation no response will return only status will return for http request
                        payload.isNoJsonResponse = true;
                        break;
                }
                return { headers: _headers, payload: payload };
            };
            return {
                List: List,
                ListItem: ListItem,
                Get: Get,
                getRequestDigest: getRequestDigest,
                Post: postWithRequestDigest,
                Update: updateWithRequestDigest,
                Delete: deleteWithRequestDigest
            };
        })();
        this.rootUrl = rootWeb;
    }
    return SPRest;
}());