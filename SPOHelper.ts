

/**
 * Console Log Short  Notation
 * @param options 
 */
const CL = console.log;

/**
* Base HttpClient which is wrapper of fetch.
* @param options 
*/
const BaseClient =(url="",{headers={},method="GET",body=undefined})=>{return fetch(url,{headers,method,body})};

const BaseRequest=(url="",{headers={},method="GET",body=undefined})=>{return BaseClient(url,{headers,method,body}).then(r=>r)}

/**
* Get JSON resposnse with Base HttpClient which is wrapper of fetch.
* @param options 
*/
const GetJson=(url="",{headers={},method="GET",body=undefined})=>{return BaseClient(url,{headers,method,body}).then(r=>r.json())}

/**
*JSON.stringify short form.
*/
const JSFY = JSON.stringify;
/**
  *JSON.parse short form.
  */
const JPS = JSON.parse;

// Reference Links:
// https://scotch.io/bar-talk/copying-objects-in-javascript
// https://stackoverflow.com/questions/35959372/property-assign-does-not-exist-on-type-objectconstructor
const _headers = {
    credentials: "include", // credentials: 'same-origin',
    Accept: "application/json; odata=nometadata",
    "Content-Type": "application/json; odata=nometadata"
}
/** 
 * Get 
 * Examples SPGet("https://tenant.sharepoint.com/sites/ABCSite/_api/Lists/getbytitle('SPO List')").then(r=>console.log(r))
 */
export const SPGet = url => {
    let headers = copyObj(_headers);
    return GetJson(url, { headers }); //BaseClient(url, { headers }).then(r => r.json());
};
const _Post = ({ url, payload = {}, hdrs = {} }) => {
    let headers = MergeObj(_headers, hdrs);//Object.assign({}, _headers, hdrs)// JPS(JSFY(_headers));
    var req = BaseClient(url, { method: "POST", body: JSFY(payload), headers });
    //Skip Conversion of Json for Update and Delete Request
    if (headers["IF-MATCH"]) return req.then(r => r);
    //Changes required in future if we change httpclient
    return req.then(r => r.json());

};

/**
  * Post Request with Request Digest
  * SPPost({url:"https://tenant.sharepoint.com/sites/ABCSite/_api/Lists",payload:{Title :"POC Doc", BaseTemplate: 101,Description: 'Created From SPOHelper' }}).then(r=>console.log(r))
   * Example  : SPPost({url:"https://tenant.sharepoint.com/sites/ABCSite/_api/Lists/getbytitle('SPO List')/items",payload:{Title :"POST test",Number:123}}).then(r=>console.log(r))
  */
export const SPPost = async ({ url, payload = {}, hdrs = {} }) => {
    let headers = copyObj(_headers)//Object.assign({}, _headers)// JPS(JSFY(_headers));
    let digest = await GetDigest(url);
    // hdrs=Object.assign({},hdrs);
    hdrs["X-RequestDigest"] = digest.FormDigestValue;
    return _Post({ url, hdrs, payload });
};

/**
  * Update Request with Request Digest
   * Example  : SPUpdate({url:"https://tenant.sharepoint.com/sites/ABCSite/_api/Lists/getbytitle('SPO List')/items(1)",payload:{Title :"POST test update",Number:1234}}).then(r=>console.log(r))
  */
export const SPUpdate = ({ url, payload = {} }) => {
    let hdrs = {
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
    };
    return SPPost({ url, hdrs, payload });
};


/**
  * Delete Request with Request Digest
  * Example  : SPDelete({url:"https://tenant.sharepoint.com/sites/ABCSite/_api/Lists/getbytitle('SPO List')/items(3)"}).then(r=>console.log(r))
  */
export const SPDelete = ({ url }) => {
    let hdrs = {
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE"
    };
    return SPPost({ url, hdrs });
};

/**
  * Get Request Digest
  */
export const GetDigest = (url = "") => {
    //// TODO -Cached based Request Digest

    //let headers = Object.assign({},_headers)// JPS(JSFY(_headers));
    // if (url) {
    //   url = this.rootUrl;
    // }
    //null, undefined check
    url = url ? url : "";
    if (typeof url != "string") return new Error("Invalid url");

    //Assuming user entered full url 
    let urlSegments = url.toLowerCase().split("/_api");
    if (urlSegments.length > 1) {
        url = `${urlSegments[0]}/_api/contextinfo`;
    } else {
        CL(url.toLowerCase().split("/_api"));
    }
    return _Post({ url });
};
const copyObj = (obj) => {
    var tempObj = { ...obj };
    //Object.assign({}, _headers)// JPS(JSFY(_headers));
    return tempObj;
}
const MergeObj = (obj1, obj2) => {
    var tempObj = { ...obj1, ...obj2 };
    //Object.assign({}, _headers)// JPS(JSFY(_headers));
    return tempObj;
}