# Sharepoint 2013/2016/2019/Online, Office 365 REST API Code Sample/Example
In this repo contains Sharepoint 2013/2016/2019/Online, Office 365 REST API Code Sample/Example which will using SP Rest utility `SPRest.ts` or `SPRest.js`
Here utility library can be used with TypeScript in #Spfx and also work with most browsers.

As I have used `fetch` API which is not available in IE11 browser so you can use [polyfill](https://github.com/github/fetch)

## Sharepoint 2013/2016/2019/Online List Item CRUD Operation

```js
var spRest=new SPRest("https://brgrp.sharepoint.com");

// Get All Item from List item  
spRest.Utils.ListItem.GetAllItem({listName:"PlaceHolderList"}).then(function(r){  
console.log(r);  
// Response received. TODO bind record to table or somewhere else.  
});

//Get all selected column Data using full URL  
var reqUItemUrl="https://brgrp.sharepoint.com/_api/web/lists/getbytitle('PlaceHolderList')/items";
util.Utils.ListItem  
.GetAllItem({  
   url:reqUItemUrl+"?$select=Id,Title&$top=200"  
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
util.Utils.ListItem.GetItemById({listName:"PlaceHolderList",Id:201}).then(function(r){    
console.log(r);    
// Response received.   
});

// Add ListItem to Sharepoint List  
util.Utils.ListItem.Add({listName:"PlaceHolderList"
,data:{Title:"New Item Created For Demo",UserId:1,Completed:"true"}}).then(function(r){    
console.log(r);    
// Added New List item response received with newly created item  
}); 

// Update List item based on ID with new data in SharePoint List  
util.Utils.ListItem.Update({listName:"PlaceHolderList",Id:201
,data:{Title:"Updated List Item",UserId:1,Completed:"true"}}).then(function(r){    
// List Item Updated and received response with status 204  
console.log(r);  
}); 

// Delete List item based on ID  
util.Utils.ListItem.Delete({listName:"PlaceHolderList",Id:201}).then(function(r){    
// List Item Deleted and received response with status 200  
console.log(r);  
}); 
```

Reference link : https://www.c-sharpcorner.com/article/easy-sharepoint-listitem-crud-operation-using-rest-api-wrapper/

## Utility method for getting RequestDigest for POST Request

```js
var getRequestDigest=(rootUrl)=>{
var _payloadOptions = {  method: "POST", 
                headers: {  credentials: "include",  Accept: "application/json; odata=verbose"
                ,"Content-Type": "application/json; odata=verbose" }  
            };  
  
//RequestDigest Request
return fetch(rootUrl+"/_api/contextinfo",_payloadOptions).then(r=>r.json())
}
```

## Upload or Create txt file in SharePoint Document Library

```js
//Get Digest first then create txt file
getRequestDigest("https://brgrp.sharepoint.com").then (r=>{
//Received Request Digest
var reqUrl="https://brgrp.sharepoint.com/_api/web/GetFolderByServerRelativeUrl('/Shared Documents')"
fetch(reqUrl+"/Files/add(url='file_name.txt',overwrite=true)",
{method:"POST",headers:
{accept:"application/json;odata=verbose",
"Content-Type":"application/json;odata=verbose","X-RequestDigest":r.d.GetContextWebInformation.FormDigestValue }
,body:"Content Of Text File"}).then(r=>console.log(r))
})
```

## Update A SharePoint List Item Without Increasing Its Item File Version Using Rest API


```js
//payload for request   
 body=  {"formValues":[{"FieldName":"Title","FieldValue":"Single Update Title with versioning__"}]
 ,bNewDocumentUpdate:true}  
  
 //Header data for sharepoint POST Request  
 var _payloadOptions = {  
                method: "POST",  
                body: undefined,  
                headers: {  
                    credentials: "include",  
                    Accept: "application/json; odata=verbose",  
                    "Content-Type": "application/json; odata=verbose"  
                }  
            };  
  
//Get RequestDigest First  
fetch("https://brgrp.sharepoint.com/_api/contextinfo",_payloadOptions).then(r=>r.json())  
.then(r=>  
                                   {  
_payloadOptions.headers["X-RequestDigest"]=r.d.GetContextWebInformation.FormDigestValue  
      
_payloadOptions.body=JSON.stringify(body);  
  
// Make REST API Call to update list item without increamenting version.  
fetch("https://brgrp.sharepoint.com/_api/web/Lists/GetbyTitle('Documents')/items(1)/ValidateUpdateListItem()",
_payloadOptions).then(r=>r.json()).then(r=>console.log(r))  
});

```

Reference link : 
https://www.c-sharpcorner.com/article/update-a-sharepoint-list-item-without-increasing-its-item-file-version-using-res/

## Make Batch Request call in SharePoint Online using BatchUtils for all Get Operation

BatchUtils can be found in [Here](https://github.com/anomepani/sp-rest-util/blob/master/BatchUtils.ts)

Here rootUrl required to Generate Request Digest Token as batch Request is POST request.

```js
var arr=["https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(212)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(213)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(214)"];

BatchUtils.GetBatchAll({rootUrl:"https://brgrp.sharepoint.com",
batchUrls:arr}).then(r=>console.log(r))

```
You can skip rootUrl if you have already generated request digest as below.

```js
var arr=["https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(212)"
, "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(213)"
, "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(214)"];

getRequestDigest("https://brgrp.sharepoint.com").then(r=>{

BatchUtils.GetBatchAll({rootUrl:"https://brgrp.sharepoint.com",
batchUrls:arr,FormDigestValue: r.d.GetContextWebInformation.FormDigestValue}).then(r=>console.log(r))
});

```

