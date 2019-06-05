# sp-rest-util
This Repo created for learning purpose to request Easily in Sharepoint REST API

I have upgraded Rest utils with name `SPRest.ts` or `SPRest.js` which covers CRUD with List Item easily.

Usage of  `SPRest.ts` or `SPRest.js` is described on [C-SharpCorner Article](https://www.c-sharpcorner.com/article/easy-sharepoint-listitem-crud-operation-using-rest-api-wrapper/)

## Below Configuration and usage for Util `SPRest.ts` or `SPRest.js`

```js
var spRest=new SPRest("https://brgrp.sharepoint.com");

// Get All Item from List item  
spRest.Utils.ListItem.GetAllItem({listName:"PlaceHolderList"}).then(function(r){  
console.log(r);  
// Response received. TODO bind record to table or somewhere else.  
});

//Get all selected column Data using full URL  
util.Utils.ListItem  
.GetAllItem({  
   url:"https://brgrp.sharepoint.com/_api/web/lists/getbytitle('PlaceHolderList')/items?$select=Id,Title&$top=200"  
}).then(function(r){console.log(r);  
//Response Received  
});  
  
//Get all selected column data with listName and oDataOption  
  
util.Utils.ListItem.GetAllItem(  
   {"listName":"PlaceHolderList"  
      ,oDataOption:"$select=Id,Title&$top=200"  
   }).then(function(r){console.log(r);  
   //Response Received  
});  

// Get List Item By Id  
spRest.Utils.ListItem.GetItemById({listName:"PlaceHolderList",Id:201}).then(function(r){    
console.log(r);    
// Response received.   
});

// Add ListItem to Sharepoint List  
spRest.Utils.ListItem.Add({listName:"PlaceHolderList",data:{Title:"New Item Created For Demo",UserId:1,Completed:"true"}}).then(function(r){    
console.log(r);    
// Added New List item response received with newly created item  
}); 

// Update List item based on ID with new data in SharePoint List  
spRest.Utils.ListItem.Update({listName:"PlaceHolderList",Id:201,data:{Title:"Updated List Item",UserId:1,Completed:"true"}}).then(function(r){    
// List Item Updated and received response with status 204  
console.log(r);  
}); 

// Delete List item based on ID  
spRest.Utils.ListItem.Delete({listName:"PlaceHolderList",Id:201}).then(function(r){    
// List Item Deleted and received response with status 200  
console.log(r);  
}); 

```

## Below Configuration and usage for Old Util  `spRestUtils.js`
```js

//Troubleshoots Possible error and its resolution

// body is not in stringify => "Invalid JSON. A token was not recognized in the JSON content."
// Request URI is not exist=> "Cannot find resource for the request list."
// While Create or Update Operation List Title,Field is already exist with name => "A list, survey, discussion board, or document library with the specified title already exists in this Web site.  Please choose another title."
// Failed to execute 'fetch' on 'Window': Request with GET/HEAD method cannot have body =>  This was due to this utility method, I have passed body in GET Request which is not supported.
// the type SP.ListCollection does not support HTTP PATCH method => While Creating new List/List Item if Http-Method is "Merge" this error occurs

// Errror : An unexpected 'StartObject' node was found when reading from the JSON reader. A 'PrimitiveValue' node was expected.

// solution: If Body have json payload then first apply stringify on that data and then again stringify whole body

// Column 'title' does not exist. It may have been deleted by another user.

// Solution: column is case sensetive in Sharepoint pass exact Column name meaning "Title"



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
