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
    "ListItem": { "__metadata": { "type": "SP.Data.AWSResponseListListItem" }, "Title": "Test" },
}
var msgraphUtils = (() => { 
    let isOdataVerbose = true;
    let _rootUrl = '';
    //Here "Accept" and "Content-Type" header must require for Sharepoint Onlin REST API
    let _payloadOptions = {
            method: 'GET',
            headers: { "Accept": "application/json; odata=verbose", "Content-Type": "application/json; odata=verbose" }
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
    const spFetchGet = (url, payload) => {
        _payloadOptions.method = "GET";
        //Internally if body is set for GET request then need to remove it by setting undefined
        _payloadOptions.body = undefined;
        console.log(_payloadOptions, payload);
        return fetch(url, _payloadOptions).then(r => r.json());
    };
    const getFetchRequestDigest = (url, payload) => {
        _payloadOptions.method = "POST";
        return fetch(url, _payloadOptions).then(r => r.json());
    };
    const spFetchPostWithDigest = (url, payload) => {
        return getFetchRequestDigest(payload.rootUrl + "/_api/contextinfo").then(token => {
            payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
            return spFetchPost(url, payload);
        })
    }
    const spFetchUpdateWithDigest = (url, payload) => {
        return getFetchRequestDigest(payload.rootUrl + "/_api/contextinfo").then(token => {
            payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
            paylod._extraHeaders = {
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE"
            };
            //For Update operation or Merge Operation no response will return only status will return for http request
            paylod.isNoJsonResponse = true;
            return spFetchPost(url, payload);
        })
    }
    const spFetchPost = (url, payload) => {
        // TODO For Safety this method can be wrapped with request Digest so always get token.
        // But need to ensure it request only when request digest is expired.
        _payloadOptions.method = "POST";
        _payloadOptions.body = payload.data;
        let _metaInfo = payload.metaInfo;
        //Pre validation Check Before update body or meta detail
        if (_metaInfo && spDefaultMeta[_metaInfo.type]) {
            //Update Title and Description while creating new List/Column/Fields
            let { type, title,listName } = _metaInfo;
            let _body = spDefaultMeta[type];
            _body.Title = title;

            if(type==="ListItem"){
                //TODO Extra Efforts
  
                _body.__metadata.type=getItemTypeForListName(listName);
            }
            //Pass body data as stringyfy;
            _payloadOptions.body = JSON.stringify(_body);
            _payloadOptions.body.Title = title;

        }
        //If Extra header is present in payload for update or other operation, Append to existing header
        if (payload._extraHeaders) {
            for (var _header in payload._extraHeaders) {
                _payloadOptions.headers[_header] = payload._extraHeaders[key];
            }
        }
        
        _payloadOptions.headers["X-RequestDigest"] = payload.requestDigest;

        console.log(_payloadOptions);
        //TODO- Naming convention can be updated.
        if (paylod.isNoJsonResponse) {
            return fetch(url, _payloadOptions).then(r => r);
        } else {
            return fetch(url, _payloadOptions).then(r => r.json());
        }

    };
    const spFetchPut = (url, payload) => {
        _payloadOptions.method = "PUT";
        _payloadOptions.body = payload.data;
        return fetch(url, _payloadOptions).then(r => r.json());
    };
    const spFetchDelete = (url, payload) => {
        _payloadOptions.method = "DELETE";
        _payloadOptions.body = payload.data;
        return fetch(url, _payloadOptions).then(r => r.json());
    };
    return { spFetchGet, spFetchPost, getFetchRequestDigest, spFetchPostWithDigest, spFetchUpdateWithDigest }
})();

//Possible error and its resolution

// body is not in stringify => "Invalid JSON. A token was not recognized in the JSON content."
// Request URI is not exist=> "Cannot find resource for the request list."
// While Create or Update Operation List Title,Field is already exist with name => "A list, survey, discussion board, or document library with the specified title already exists in this Web site.  Please choose another title."
// Failed to execute 'fetch' on 'Window': Request with GET/HEAD method cannot have body =>  This was due to this utility method, I have passed body in GET Request which is not supported.

// How to use this msgraphUtils wrapper method to make sharepoint Online REST API Easy and simple
// 1. To Make a Post Request e.g. Create List in Sharepoint then pass paramater as below
//rootUrl="https://brgroup.sharepoint.com"
// requestUrl=rootUrl+"/_api/web/lists"
//Here "type" is meta type which internally fetched and passed in request and "title" is name of list title you want while creation
// metaInfo={type:"List",title:"List Created with minimal info"},
//requestDigest:temp2.d.GetContextWebInformation.FormDigestValue
// ;msgraphUtils.spFetchPost(requestUrl,{metaInfo:metaInfo,requestDigest:requestDigest}).then(r=>console.log(r))
// 2.  Request  for digest value
// Internally we have set headers and method :"POST" so you have to just pass only request url.
// requestUrl=rootUrl+"/_api/contextinfo"
//msgraphUtils.getFetchRequestDigest(requestUrl).then(r=>console.log(r))
// 3. Make Post request with  inbuilt Request Digest in chaining.
// rootUrl="https://brgroup.sharepoint.com"
// requestUrl=rootUrl+"/_api/web/lists"
//Here "type" is meta type which internally fetched and passed in request and "title" is name of list title you want while creation
// metaInfo={type:"List",title:"List Created with minimal info"},
// In payload for below request I have passed extra param "rootUrl" which require internally.
// msgraphUtils.spFetchPostWithDigest(rootUrl+"/_api/web/lists",{rootUrl:rootUrl,metaInfo:{type:"List",title:"List Created with msgraphutils 0.11 with digest"}}).then(r=>console.log(r))
// Internally we have set headers and method :"POST" so you have to just pass only request url.
// requestUrl=rootUrl+"/_api/contextinfo"
//msgraphUtils.getFetchRequestDigest(requestUrl).then(r=>console.log(r))

//Working Sample Full Request
//msgraphUtils.spFetchGet(rootUrl+"/_api/web/lists").then(r=>console.log(r))
//msgraphUtils.getFetchRequestDigest(requestUrl).then(r=>console.log(r))
//msgraphUtils.spFetchPost(rootUrl+"/_api/web/lists",{metaInfo:{type:"List",title:"List Created with msgraphutils 0.1"},requestDigest:temp5.d.GetContextWebInformation.FormDigestValue}).then(r=>console.log(r))
//msgraphUtils.spFetchPostWithDigest(rootUrl+"/_api/web/lists",{rootUrl:rootUrl,metaInfo:{type:"List",title:"List Created with msgraphutils 0.11 with digest"}}).then(r=>console.log(r))
//reqUrl="https://brgroup.sharepoint.com/_api/web/lists/getbytitle('AWSResponseList')/items"
//msgraphUtils.spFetchGet(reqUrl+"?$filter=Title eq 'Test Response'").then(r=>console.log(r))
//msgraphUtils.spFetchGet(reqUrl+"?$filter=Title eq 'Test Response' and TargetLanguage eq 'pt-br'").then(r=>console.log(r))
//msgraphUtils.spFetchGet(reqUrl+"?$filter=Title eq 'Test Response' and TargetLanguage eq 'pt-br'").then(r=>console.log(r))