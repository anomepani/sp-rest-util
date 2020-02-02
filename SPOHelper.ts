

/**
 * Console Log Short  Notation
 * @param options 
 */
const CL = console.log;

/**
 * Check isObject type is Json
 * Reference from : https://stackoverflow.com/questions/11182924/how-to-check-if-javascript-object-is-json
 * */
const isObject = (obj) => obj !== undefined && obj !== null && obj.constructor == Object;
/**
* Base HttpClient which is wrapper of fetch.
* @param options 
*/
const BaseClient = (url = "", { headers = {}, method = "GET", body = undefined }) => { return fetch(url, { headers, method, body }); };

const BaseRequest = (url = "", { headers = {}, method = "GET", body = undefined }) => { return BaseClient(url, { headers, method, body }).then(r => r); };

/**
* Get JSON resposnse with Base HttpClient which is wrapper of fetch.
* @param options 
*/
const GetJson = (url = "", { headers = {}, method = "GET", body = undefined }) => { return BaseClient(url, { headers, method, body }).then(r => r.json()); };

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
};

let defaultDigest = { FormDigestValue: "" };
const copyObj = (obj) => {
    var tempObj = { ...obj };
    //Object.assign({}, _headers)// JPS(JSFY(_headers));
    return tempObj;
};
const MergeObj = (obj1, obj2) => {
    var tempObj = { ...obj1, ...obj2 };
    //Object.assign({}, _headers)// JPS(JSFY(_headers));
    return tempObj;
};

const _Post = ({ url, payload = {}, hdrs = {}, isBlobOrArrayBuffer = false }) => {
    let headers = MergeObj(_headers, hdrs);//Object.assign({}, _headers, hdrs)// JPS(JSFY(_headers));
    var req = BaseClient(url, { method: "POST", body: isBlobOrArrayBuffer ? payload : JSFY(payload), headers });
    //Skip Conversion of Json for Update and Delete Request
    if (headers["IF-MATCH"]) return req.then(r => r);
    //Changes required in future if we change httpclient
    return req.then(r => r.json());

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

/** 
 * Get 
 * Examples SPGet("https://tenant.sharepoint.com/sites/ABCSite/_api/Lists/getbytitle('SPO List')").then(r=>console.log(r))
 */
 export const SPGet = url => {
    let headers = copyObj(_headers);
    return GetJson(url, { headers }); //BaseClient(url, { headers }).then(r => r.json());
};

/**
  * Post Request with Request Digest
  * SPPost({url:"https://tenant.sharepoint.com/sites/ABCSite/_api/Lists",payload:{Title :"POC Doc", BaseTemplate: 101,Description: 'Created From SPOHelper' }}).then(r=>console.log(r))
   * Example  : SPPost({url:"https://tenant.sharepoint.com/sites/ABCSite/_api/Lists/getbytitle('SPO List')/items",payload:{Title :"POST test",Number:123}}).then(r=>console.log(r))
  */
 export const SPPost = async ({ url, payload = {}, hdrs = {}, digest = defaultDigest, isBlobOrArrayBuffer = false }) => {
    //If Digest is null or undefined then request for new digest otherwise pass digest as it is
    ///NEED TO TEST Below logic

    if (!digest || (isObject(digest) && !digest.FormDigestValue))
        digest = await GetDigest(url);

    hdrs["X-RequestDigest"] = digest.FormDigestValue;
    return _Post({ url, hdrs, payload, isBlobOrArrayBuffer });
};

/**
  * Add attachment using SPFileUpload metod with payload Blob type
  * SPFileUpload({url:"https://tenant.sharepoint.com/_api/Lists/GetByTitle('SPOList')/items(1)//AttachmentFiles/ add(FileName='abc3.txt') ",payload:new Blob(["This is text"],{type:"text/plain"})}).then(r=>console.log(r))
  */
 export const SPFileUpload = async ({ url, payload = {}, hdrs = {}, digest = defaultDigest }) => {
    return SPPost({ isBlobOrArrayBuffer: true, url, hdrs, digest, payload });
};

 export const SPMultiFileUpload =  ({ url, payload = {}, hdrs = {}, files = [], digest = defaultDigest }) => {
var promiseAll =[];
//Wait Till all file uploading.
return new Promise(async(resolve,reject)=>{


  //  files.forEach(async (i,v) =>
//console.log("Processing iles",files.length);
for(var j=0;j<files.length;j++)
 {
var i=files[j]
//console.log("Processing",j,i);
//Uploading file one by one in loop 
        //TODO Encode/decode filename
        var abc= await SPFileUpload(
            {
                url: `${url}add(FileName='${i.fileName}')`
                , payload: i.data, digest
            });


promiseAll.push(abc);
if(promiseAll.length==files.length)
resolve(promiseAll);
    };
});
};


/**
  * Update Request with Request Digest
   * Example  : SPUpdate({url:"https://tenant.sharepoint.com/sites/ABCSite/_api/Lists/getbytitle('SPO List')/items(1)",payload:{Title :"POST test update",Number:1234}}).then(r=>console.log(r))
  */
 export const SPUpdate = ({ url, payload = {}, digest = defaultDigest }) => {
    let hdrs = {
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
    };
    return SPPost({ url, hdrs, payload, digest });
};
/**
  * Delete Request with Request Digest
  * Example  : SPDelete("https://tenant.sharepoint.com/sites/ABCSite/_api/Lists/getbytitle('SPO List')/items(3)").then(r=>console.log(r))
  */
 export const SPDelete = (url, digest = defaultDigest) => {
    let hdrs = {
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE"
    };
    return SPPost({ url, hdrs, digest });
};

let GetUniqueFileData=()=>{return new Blob([GetUniqueFileName()])}

let GetUniqueFileName=()=>new Date().getTime() +".txt";

let GenerateSampleFileArray=(count)=>{
return Array.from(Array(count)).map((_, i) => { 
//console.log(_,i)
var fileName=GetUniqueFileName(),data=GetUniqueFileData();
return {fileName,data} })};