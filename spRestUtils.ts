class SPRest{
    rootUrl: any;
    constructor(rootWeb){
        this.rootUrl=rootWeb;
    }

 Utils = (() => {
    /**
 * Reference or motivation link : https://github.com/omkarkhair/sp-jello/blob/master/lib/jello.js
 */
var reqUrl = this.rootUrl + "/_api/web/lists/getbytitle('AWSResponseList')/items";
//var fetch = require("node-fetch");
// TODO- Need to transform this json to closure function to allow only get operation.
//TODO  - Need to add support for Field Meta in utils
var FieldTypeKind = [{ "Invalid": 0 }, { "Integer": 1 }, { "Text": 2 }, { "Note": 3 }, { "DateTime": 4 }, { "Counter": 5 }, { "Choice": 6 }, { "Lookup": 7 }, { "Boolean": 8 }, { "Number": 9 }, { "Currency": 10 }, { "URL": 11 }, { "Computed": 12 }, { "Threading": 13 }, { "Guid": 14 }, { "MultiChoice": 15 }, { "GridChoice": 16 }, { "Calculated": 17 }, { "File": 18 }, { "Attachments": 19 }, { "User": 20 }, { "Recurrence": 21 }, { "CrossProjectLink": 22 }, { "ModStat": 23 }, { "Error": 24 }, { "ContentTypeId": 25 }, { "PageSeparator": 26 }, { "ThreadIndex": 27 }, { "WorkflowStatus": 28 }, { "AllDayEvent": 29 }, { "WorkflowEventType": 30 }, { "Geolocation": "" }, { "OutcomeChoice": "" }, { "MaxItems": 31 }];
var spDefaultMeta = {
    "List": {
        BaseTemplate: 100,
        // "Title": "List Created with Fetch UTIL1112",
        "__metadata": { type: "SP.List" }
    },
    "Document": {
        '__metadata': { 'type': 'SP.List' },
        'AllowContentTypes': true,
        'BaseTemplate': 101,
        'ContentTypesEnabled': true,
        // 'Description': 'My doc. lib. description',
        // 'Title': 'Test'
    },
    "ListItem": {
        "__metadata": { "type": "SP.Data.AWSResponseListListItem" },
        // "Title": "Test"
    },
};

    console.log(this.rootUrl);
    let isOdataVerbose = true;
    let _rootUrl = this.rootUrl;
    //Here "Accept" and "Content-Type" header must require for Sharepoint Onlin REST API
    let _payloadOptions = {
        method: 'GET',
        headers: { "Accept": "application/json; odata=verbose", "Content-Type": "application/json; odata=verbose" }
        , body: undefined
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
        console.log(url);
        _payloadOptions.method = "GET";
        //Internally if body is set for GET request then need to remove it by setting undefined
        // otherwise it return with error : Failed to execute 'fetch' on 'Window': Request with GET/HEAD method cannot have body
        _payloadOptions.body = undefined;
        console.log(_payloadOptions);
        return fetch(url, _payloadOptions).then(r => r.json());
    };

    const getRequestDigest = (url = '') => {
        if (url) {
            url = _rootUrl;
        }
        url += "/_api/contextinfo";
        _payloadOptions.method = "POST";
        return fetch(url, _payloadOptions).then(r => r.json());
    };
    const postWithRequestDigestExtension = (url, { headers, payload }) => {
        return getRequestDigest().then(token => {
            // payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
            headers.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
            return PostExtension(url, { headers, payload });
        })
    }
    const postWithRequestDigest = (url, payload) => {
        return getRequestDigest().then(token => {
            payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
            return Post(url, payload);
        })
    }
    //Need to refactor below method and need to merge in single method postWithRequestDigestExtension
    //Based on Action Type in headers using switch case add or remove extra headers
    // Try to add header name same as standard name so it can be replaced with for loop e.g. _extraHeaders
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
    const deleteWithRequestDigest = (url, payload) => {
        return getRequestDigest(payload.rootUrl + "/_api/contextinfo").then(token => {
            payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
            payload._extraHeaders = {
                "IF-MATCH": "*",
                "X-HTTP-Method": "DELETE"
            };
            //For Update operation or Merge Operation no response will return only status will return for http request
            payload.isNoJsonResponse = true;
            return Post(url, payload);
        })
    }


    const Post = (url, payload) => {
        // TODO For Safety this method can be wrapped with request Digest so always get token.
        // But need to ensure it request only when request digest is expired.
        _payloadOptions.method = "POST";
        _payloadOptions.body = payload.data;
        let _metaInfo = payload.metaInfo;
        //Pre validation Check Before update body or meta detail
        if (_metaInfo && spDefaultMeta[_metaInfo.type]) {
            //Update Title and Description while creating new List/Column/Fields
            let { type, title, listName } = _metaInfo;
            let _body = spDefaultMeta[type];
            _body.Title = title;

            if (type === "ListItem") {
                //TODO Extra Efforts

                _body.__metadata.type = getItemTypeForListName(listName);
            }
            //Pass body data as stringyfy;
            _payloadOptions.body = JSON.stringify(_body);
            _payloadOptions.body.Title = title;

        }
        //If Extra header is present in payload for update or other operation, Append to existing header
        if (payload._extraHeaders) {
            for (var _header in payload._extraHeaders) {
                _payloadOptions.headers[_header] = payload._extraHeaders[_header];
            }
        } else {
            //IF Not present that means it is POST Request, reset Extraheader for this request

            _payloadOptions.headers["IF-MATCH"] = undefined;
            _payloadOptions.headers["X-HTTP-Method"] = undefined;
        }

        _payloadOptions.headers["X-RequestDigest"] = payload.requestDigest;

        console.log(_payloadOptions);
        //TODO- Naming convention can be updated.
        if (payload.isNoJsonResponse) {
            return fetch(url, _payloadOptions).then(r => r);
        } else {
            return fetch(url, _payloadOptions).then(r => r.json());
        }

    };
    const PostExtension = (url, {headers,payload}) => {
        // TODO For Safety this method can be wrapped with request Digest so always get token.
        // But need to ensure it request only when request digest is expired.
        _payloadOptions.method = "POST";
        _payloadOptions.headers["X-RequestDigest"] = headers.requestDigest;
        _payloadOptions.body = payload.data;
        // let _metaInfo = payload.metaInfo;
        // //Pre validation Check Before update body or meta detail
        // // if (_metaInfo && spDefaultMeta[_metaInfo.type]) {
        // //     //Update Title and Description while creating new List/Column/Fields
        // //     let { type, title, listName } = _metaInfo;
        // //     let _body = spDefaultMeta[type];
        // //     _body.Title = title;

        // //     if (type === "ListItem") {
        // //         //TODO Extra Efforts

        // //         _body.__metadata.type = getItemTypeForListName(listName);
        // //     }
        // //     //Pass body data as stringyfy;
        // //     _payloadOptions.body = JSON.stringify(_body);
        // //     _payloadOptions.body.Title = title;

        // // }
        // // //If Extra header is present in payload for update or other operation, Append to existing header
        // // if (payload._extraHeaders) {
        // //     for (var _header in payload._extraHeaders) {
        // //         _payloadOptions.headers[_header] = payload._extraHeaders[_header];
        // //     }
        // // } else {
        // //     //IF Not present that means it is POST Request, reset Extraheader for this request

        // //     _payloadOptions.headers["IF-MATCH"] = undefined;
        // //     _payloadOptions.headers["X-HTTP-Method"] = undefined;
        // // }

        // _payloadOptions.headers["X-RequestDigest"] = payload.requestDigest;

        // console.log(_payloadOptions);
        //TODO- Naming convention can be updated.
       // if (payload.isNoJsonResponse) {
           //Instead of storing  extra info in payload use headers
           if (headers.isNoJsonResponse) {
            return fetch(url, _payloadOptions).then(r => r);
        } else {
            return fetch(url, _payloadOptions).then(r => r.json());
        }

    };
    const Put = (url, payload) => {
        _payloadOptions.method = "PUT";
        _payloadOptions.body = payload.data;
        return fetch(url, _payloadOptions).then(r => r.json());
    };
    const Delete = (url, payload) => {
        _payloadOptions.method = "DELETE";
        _payloadOptions.body = payload.data;
        return fetch(url, _payloadOptions).then(r => r.json());
    };

    const generateRestRequest = ({ listName, Id = '' }) => {
        let prepareRequest = `${_rootUrl}/_api/web/lists/getbytitle('${listName}')/items`;
        if (Id) {
            prepareRequest += `/${Id}`;
        }

        return prepareRequest;
    };
//Provide support for odata query so, for specific provided expression append it with base requet
    const ListItem = {
        Add: ({ }: any) => { },
        Update: () => { },
        Delete: () => { },
        GetItemById: ({ }: any) => { },
        GetAllItem: ({ }: any) => { },
    };

    ListItem.GetAllItem = ({ listName }) => {
        return Get(generateRestRequest({ listName }));
    };

    ListItem.GetItemById = ({ listName, Id }) => {
        return Get(generateRestRequest({ listName, Id }));
    };
     // Add/Update/Delete require same types of metaData as well as payload only difference is Headers

    ListItem.Add = ({ listName, data }) => {
        var _reqUrl = generateRestRequest({ listName });
        var payload: any = {};
        let _metaInfo: any = spDefaultMeta["ListItem"];
        _metaInfo.type = getItemTypeForListName(listName);
        payload.__metadata = _metaInfo;
        //Need to check is data is object then iterate otherwise not
        if (data) {
            for (let _key in data) {
                payload[_key] = data[_key];
            }
        }
        //TODO Verify about extra header

        // //Pass body data as stringyfy;
        // // _payloadOptions.body = JSON.stringify(_body);
        // // _payloadOptions.body.Title = title;

        //    return postWithRequestDigest(_reqUrl);

        return postWithRequestDigestExtension(_reqUrl, payload);

    }
   
    return { ListItem: ListItem, Get: Get, getRequestDigest: getRequestDigest, Post: postWithRequestDigest, Update: updateWithRequestDigest, Delete: deleteWithRequestDigest }
})();
}