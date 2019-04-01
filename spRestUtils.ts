/**
 * Reference or motivation link : https://github.com/omkarkhair/sp-jello/blob/master/lib/jello.js
 */

//var fetch = require("node-fetch");
// TODO- Need to transform this json to closure function to allow only get operation.
//TODO  - Need to add support for Field Meta in utils
var FieldTypeKind = [{ "Invalid": 0 }, { "Integer": 1 }, { "Text": 2 }, { "Note": 3 }, { "DateTime": 4 }, { "Counter": 5 }, { "Choice": 6 }, { "Lookup": 7 }, { "Boolean": 8 }, { "Number": 9 }, { "Currency": 10 }, { "URL": 11 }, { "Computed": 12 }, { "Threading": 13 }, { "Guid": 14 }, { "MultiChoice": 15 }, { "GridChoice": 16 }, { "Calculated": 17 }, { "File": 18 }, { "Attachments": 19 }, { "User": 20 }, { "Recurrence": 21 }, { "CrossProjectLink": 22 }, { "ModStat": 23 }, { "Error": 24 }, { "ContentTypeId": 25 }, { "PageSeparator": 26 }, { "ThreadIndex": 27 }, { "WorkflowStatus": 28 }, { "AllDayEvent": 29 }, { "WorkflowEventType": 30 }, { "Geolocation": "" }, { "OutcomeChoice": "" }, { "MaxItems": 31 }];
var spDefaultMeta = {
    "List": {
        BaseTemplate: 100,
        "Title": "List Created with Fetch UTIL1112",
        "__metadata": { type: "SP.List" }
    },
    "Document": {
        '__metadata': { 'type': 'SP.List' },
        'AllowContentTypes': true,
        'BaseTemplate': 101,
        'ContentTypesEnabled': true,
        'Description': 'My doc. lib. description',
        'Title': 'Test'
    },
    "SitePagesItem":{
    "__metadata":{
    "type": "SP.Data.SitePagesItem"
}},

    "ListItem": { "__metadata": { "type": "SP.Data.AWSResponseListListItem" }, "Title": "Test" },
}
const spRestUtils= (() => {
    let isOdataVerbose = true;
    let _rootUrl = '';
    //Here "Accept" and "Content-Type" header must require for Sharepoint Onlin REST API
    let _payloadOptions = {
            method: 'GET',
            body:undefined,
            headers: { credentials: "include","Accept": "application/json; odata=verbose", "Content-Type": "application/json; odata=verbose" }
        }
        // TODO- To support for param base odata
        //No to fetch Metadata, response only requested data.
        // Reference link https://www.microsoft.com/en-us/microsoft-365/blog/2014/08/13/json-light-support-rest-sharepoint-api-released/
        //Option 1: verbose “accept: application/json; odata=verbose”
        //Option 2: minimalmetadata “accept: application/json; odata=minimalmetadata”
        //Option 3: nometadata “accept: application/json; odata=nometadata”
        //Option 4: Don’t provide it “accept: application/json” This defaults to minimalmetadata option
    if (!isOdataVerbose) {
        _payloadOptions.headers.Accept = "application/json; odata=nometadata"
    }
    // Reference from :https://www.tjvantoll.com/2015/09/13/fetch-and-errors/

    //Reference rom : https://sharepoint.stackexchange.com/questions/105380/adding-new-list-item-using-rest

    // Get List Item Type metadata
    const getItemTypeForListName = (name) => {
        return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
    }
    const Get = (url) => {
        var _localPayload = _payloadOptions;
        _localPayload.method = "GET";
        //Internally if body is set for GET request then need to remove it by setting undefined
        // otherwise it return with error : Failed to execute 'fetch' on 'Window': Request with GET/HEAD method cannot have body
        _localPayload.body = undefined;
        console.log(_localPayload);
        return fetch(url, _localPayload).then(r => r.json());
    };
    const getRequestDigest = (url) => {
        var _localPayload = _payloadOptions;
        _localPayload.method = "POST";
        return fetch(url, _localPayload).then(r => r.json());
    };
    const postWithRequestDigest = (url, payload) => {
        return getRequestDigest(payload.rootUrl + "/_api/contextinfo").then(token => {
            payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
            return Post(url, payload);
        })
    }
    const updateWithRequestDigest = (url, payload) => {
        return getRequestDigest(payload.rootUrl + "/_api/contextinfo").then(token => {
            payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
            payload._extraHeaders = {
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE"
            };
            //For Update operation or Merge Operation no response will return only status will return for http request
            payload.isNoJsonResponse = true;
            return Post(url, payload);
        })
    }
    const Post = (url, payload) => {
        var _localPayload = _payloadOptions;
        // TODO For Safety this method can be wrapped with request Digest so always get token.
        // But need to ensure it request only when request digest is expired.
        _localPayload.method = "POST";
        _localPayload.body = payload.data;
        let _metaInfo = payload.metaInfo;
        //Pre validation Check Before update body or meta detail
        if (_metaInfo && spDefaultMeta[_metaInfo.type]) {
            //Update Title and Description while creating new List/Column/Fields
            let { type, title, listName } = _metaInfo;
            let _body = spDefaultMeta[type];
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
                for (let field in payload.extraFields) {
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
            return fetch(url, _localPayload).then(r => r);
        } else {
            return fetch(url, _localPayload).then(r => r.json());
        }

    };
    const Put = (url, payload) => {
        var _localPayload = _payloadOptions;
        _localPayload.method = "PUT";
        _localPayload.body = payload.data;
        return fetch(url, _payloadOptions).then(r => r.json());
    };
    const Delete = (url, payload) => {
        var _localPayload = _payloadOptions;
        _localPayload.method = "DELETE";
        _localPayload.body = payload.data;
        return fetch(url, _payloadOptions).then(r => r.json());
    };
    return { Get: Get, Post: Post, getRequestDigest: getRequestDigest, postWithRequestDigest: postWithRequestDigest, updateWithRequestDigest: updateWithRequestDigest }
})();

export default spRestUtils;