# sp-rest-util
This Repo created for learning purpose to request Easily in Sharepoint REST API/ MSGraph API

```js
// Pre-requisite
var rootUrl="https://brgroup.sharepoint.com";
var reqUrl="https://brgroup.sharepoint.com/_api/web/lists/getbytitle('AWSResponseList')/items"
//Here "type" is meta type which internally used and passed in request for preparing sharepoint
//meta type
// and "title" is name of list title you want while creation which should be unique
// metaInfo={type:"List",title:"List Created with minimal info"},
// In payload for request "postWithRequestDigest" I have passed extra param "rootUrl" which
// require internally to get request digest.

// ===============================================================
// * FETCH ALL SHAREPOINT LIST ITEM USING REST UTIL
// ===============================================================

//Fetch All Items
spRestUtils.Get(reqUrl).then(r=>console.log(r));

//Fetch Item  by id
spRestUtils.Get(reqUrl+"(1)").then(r=>console.log(r))

//Fetch Item  based on filter criteria
var filterRequest= reqUrl + "?$filter=Title eq 'Test Response1' ";
spRestUtils.Get(filterRequest).then(r=>console.log(r))

// Fetch all lists from site collection
spRestUtils.Get(rootUrl+"/_api/web/lists").then(r=>console.log(r))

//Generate or Get Latest RequestDigest
spRestUtils.getRequestDigest(rootUrl+"/_api/contextinfo").then(r=>console.log(r))

// ===============================================================
// * CREATE SHAREPOINT LIST USING REST UTIL
// ===============================================================

// Create New List in Sharepoint using REST UTIL by passing request digest
spRestUtils.getRequestDigest(rootUrl+"/_api/contextinfo").then(r=>{
    //RequestDigest Received

    console.log(r);
    // Now Request new Post method

    spRestUtils.Post(rootUrl+"/_api/web/lists"
,{metaInfo:
    {type:"List",title:"List Created with spRestUtils 0.1"}
    ,requestDigest:r.d.GetContextWebInformation.FormDigestValue}
    ).then(r=>console.log("List Created",r));

});

// Create New List in Sharepoint using REST UTIL with internal request digest request

    spRestUtils.postWithRequestDigest(rootUrl+"/_api/web/lists",
    {rootUrl:rootUrl,metaInfo:{type:"List",title:"List Created with spRestUtils 0.11 with digest"}
}).then(r=>console.log(r))

// ===============================================================
// * CREATE/UPDATE SHAREPOINT LIST ITEM USING REST UTIL
// ===============================================================

// Working example for Creating List item, response return json object
spRestUtils.postWithRequestDigest(reqUrl
    ,{metaInfo:{type:"ListItem",title:"item-created-using-msgraph-utli",listName:"AWSResponseList"}
    ,rootUrl:rootUrl
}).then(r=>console.log(r));

//Verify Created List item by Searching with title 
spRestUtils.Get(reqUrl+"?$filter= Title eq 'item-created-using-msgraph-utli'").then(r=>console.log(r))

//To Update List item first fetch list item and grab item id and then construct reqUrl and update list item
// Return Response with status 204, ok :true

spRestUtils.updateWithRequestDigest(reqUrl+"(3)"
,{metaInfo:{type:"ListItem",title:"item-created-using-msgraph-utli-updated"
,listName:"AWSResponseList"},rootUrl:rootUrl})
.then(r=>console.log(r))

// ===============================================================
```