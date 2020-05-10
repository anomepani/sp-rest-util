//Reference from  Vardhman Despande blogs https://www.vrdmn.com/2016/06/sharepoint-online-get-userprofile.html
//Reference from  https://github.com/andrewconnell/sp-o365-rest/blob/master/SpRestBatchSample/Scripts/App.js
//var arr =["https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(212)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(213)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(214)"]
var BatchUtils = (function () {
    /**
       * Build the batch request for each property individually
       * @param endPointsToGet
       * @param boundryString
       */
    function buildBatchRequestBody(endPointsToGet, boundryString) {
        var propData = new Array();
        for (var i = 0; i < endPointsToGet.length; i++) {
            var getPropRESTUrl = endPointsToGet[i];
            propData.push('--batch_' + boundryString);
            propData.push('Content-Type: application/http');
            propData.push('Content-Transfer-Encoding: binary');
            propData.push('');
            propData.push('GET ' + getPropRESTUrl + ' HTTP/1.1');
            propData.push('Accept: application/json;odata=verbose');
            propData.push('');
        }
        return propData.join('\r\n');
    }
    function BuildChangeSetRequestBody(changeSetId, action, endpoint, items) {
        var batchContents = [];
        var item;
        // create the changeset
        for (var i = 0; i < items.length; i++) {
            item = items[i].data;
            action = items[i].action;
            //TODO -Need to test for different action
            endpoint = items[i].reqUrl;
            batchContents.push("--changeset_" + changeSetId);
            batchContents.push("Content-Type: application/http");
            batchContents.push("Content-Transfer-Encoding: binary");
            batchContents.push("");
            if (action === "UPDATE") {
                batchContents.push("PATCH " + endpoint + " HTTP/1.1");
                batchContents.push("If-Match: *");
                batchContents.push("Content-Type: application/json;odata=verbose");
                batchContents.push("");
                batchContents.push(JSON.stringify(item));
            }
            else if (action === "ADD") {
                batchContents.push("POST " + endpoint + " HTTP/1.1");
                //For Insert it will return created value in json format
                batchContents.push('Accept: application/json;odata=verbose');
                batchContents.push("Content-Type: application/json;odata=verbose");
                batchContents.push("");
                batchContents.push(JSON.stringify(item));
            }
            else if (action === "DELETE") {
                batchContents.push("DELETE " + endpoint + " HTTP/1.1");
                batchContents.push("If-Match: *");
            }
            //Commented POST request line and added code for UPDATE as well
            // batchContents.push("POST " + endpoint + " HTTP/1.1");
            // batchContents.push("Content-Type: application/json;odata=verbose");
            // batchContents.push("");
            // batchContents.push(JSON.stringify(_item));
            batchContents.push("");
            // // END changeset to create data
            // batchContents.push("--changeset_" + changeSetId + "--");
        }
        // END changeset to create data
        batchContents.push("--changeset_" + changeSetId + "--");
        // batch body
        return batchContents.join("\r\n");
    }
    var BuildChangeSetRequestHeader = function (batchGuid, changeSetId, batchBody) {
        var batchContents = [];
        // create batch for creating items
        batchContents.push("--batch_" + batchGuid);
        batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' +
            changeSetId +
            '"');
        batchContents.push("Content-Length: " + batchBody.length);
        batchContents.push("Content-Transfer-Encoding: binary");
        batchContents.push("");
        batchContents.push(batchBody);
        batchContents.push("");
        // create request in batch to get all items after all are created
        ////Commented below endpoint as we are utilizing same endpoint without orderby
        // endpoint = _this.rootUrl +
        //     "/_api/web/lists/getbytitle('" + listName + "')" +
        //     '/items?$orderby=Title';
        // batchContents.push('--batch_' + batchGuid);
        // batchContents.push('Content-Type: application/http');
        // batchContents.push('Content-Transfer-Encoding: binary');
        // batchContents.push('');
        //COmmented below lines of code as I don't need to request GET after insertion
        // batchContents.push('GET ' + endpoint + ' HTTP/1.1');
        // batchContents.push('Accept: application/json;odata=verbose');
        // batchContents.push('');
        batchContents.push("--batch_" + batchGuid + "--");
        return batchContents.join("\r\n");
    };
    /**
     * Build the batch header containing the user profile data as the batch body
     * @param userPropsBatchBody
     * @param boundryString
     */
    function buildBatchRequestHeader(userPropsBatchBody, boundryString) {
        var headerData = [];
        headerData.push('Content-Type: multipart/mixed; boundary="batch__' + boundryString + '"');
        headerData.push('Content-Length: ' + userPropsBatchBody.length);
        headerData.push('Content-Transfer-Encoding: binary');
        headerData.push('');
        headerData.push(userPropsBatchBody);
        headerData.push('');
        headerData.push('--batch_' + boundryString + '--');
        return headerData.join('\r\n');
    }
    /**
     * Parse Batch Get Response
     * @param batchResponse
     */
    function parseResponse(batchResponse) {
        //Extract the results back from the BatchResponse
        var results = grep(batchResponse.split("\r\n"), function (responseLine) {
            try {
                return responseLine.indexOf("{") != -1 && typeof JSON.parse(responseLine) == "object";
            }
            catch (ex) { /*adding the try catch loop for edge cases where the line contains a { but is not a JSON object*/ }
        }, null);
        //Convert JSON strings to JSON objects
        return results.map(function (result) {
            return JSON.parse(result);
        });
    }
    /**
     * Custom Batch Response Parser with Request status code
     * @param batchResponse
     */
    function customBatchResponseParser(batchResponse) {
        var batchResults = [];
        var lastStatusCode = "";
        batchResponse.split("\r\n").forEach(function (i) {
            //Creating parser for Http Status Code
            if (i.indexOf("HTTP/1.1 ") > -1) {
                console.log(i);
                //lastRequest=i.split("HTTP/1.1")[1].split(" ")[0]
                //alert(i.split("HTTP/1.1")[1].split(" ")[1])
                lastStatusCode = i.split("HTTP/1.1 ")[1].split(" ")[0];
                batchResults.push({
                    statusCode: lastStatusCode,
                    result: []
                });
            }
            if (lastStatusCode === "200" && i.indexOf("{") != -1) {
                try {
                    if (typeof JSON.parse(i) == "object") {
                        batchResults[batchResults.length - 1].result = JSON.parse(i);
                    }
                }
                catch (ex) { /*adding the try catch loop for edge cases where the line contains a { but is not a JSON object*/ }
            }
            else {
                //statusCodes.push({statusCode:lastRequest,result:[]})  
            }
        });
        return batchResults;
    }
    /**
     * Copied grep function from jQuery JavaScript Library v1.11.3
     * @param elems
     * @param callback
     * @param invert
     */
    var grep = function (elems, callback, invert) {
        var callbackInverse, matches = [], i = 0, length = elems.length, callbackExpect = !invert;
        // Go through the array, only saving the items
        // that pass the validator function
        for (; i < length; i++) {
            callbackInverse = !callback(elems[i], i);
            if (callbackInverse !== callbackExpect) {
                matches.push(elems[i]);
            }
        }
        return matches;
    };
    /**
     * Get Uniquie boundry string for batch request identifier
     */
    function getBoundryString() {
        return "vrd_" + Math.random().toString(36).substr(2, 9);
    }
    var makeBatchRequests = function (_a) {
        var rootUrl = _a.rootUrl, batchUrls = _a.batchUrls, FormDigestValue = _a.FormDigestValue;
        if (FormDigestValue) {
            return internalBatch({ FormDigestValue: FormDigestValue, rootUrl: rootUrl, batchUrls: batchUrls });
        }
        else {
            return fetch(rootUrl + "/_api/contextinfo", {
                method: "POST",
                "headers": { "Accept": "application/json;odata=verbose" },
                credentials: "include"
            }).then(function (r) { return r.json(); }).then(function (r) {
                return internalBatch({ FormDigestValue: r.d.GetContextWebInformation.FormDigestValue, rootUrl: rootUrl, batchUrls: batchUrls });
            });
        }
    };
    var makePostBatchRequests = function (_a) {
        var rootUrl = _a.rootUrl, batchUrls = _a.batchUrls, FormDigestValue = _a.FormDigestValue;
        if (FormDigestValue) {
            return internalPostBatch({ FormDigestValue: FormDigestValue, rootUrl: rootUrl, batchUrls: batchUrls });
        }
        else {
            return fetch(rootUrl + "/_api/contextinfo", {
                method: "POST",
                "headers": { "Accept": "application/json;odata=verbose" },
                credentials: "include"
            }).then(function (r) { return r.json(); }).then(function (r) {
                return internalPostBatch({ FormDigestValue: r.d.GetContextWebInformation.FormDigestValue, rootUrl: rootUrl, batchUrls: batchUrls });
            });
        }
    };
    var internalPostBatch = function (_a) {
        //Reference from  Vardhman Despande blogs https://www.vrdmn.com/2016/06/sharepoint-online-get-userprofile.html
        var FormDigestValue = _a.FormDigestValue, rootUrl = _a.rootUrl, batchUrls = _a.batchUrls;
        //AccountName of the user
        //  var userAccountName = encodeURIComponent("i:0#.f|membership|user@yourtenant.onmicrosoft.com");
        //Collection of Endpoint url to fetch
        var endpointsArray = batchUrls; //["https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(212)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(213)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(214)"]
        //Unique identifier, will be used to delimit different parts of the request body
        var boundryString = getBoundryString();
        //Unique identifier, will be used to delimit different parts of the request body
        var changeSetIdString = getBoundryString();
        //Build Body of the Batch Request
        var userPropertiesBatchBody = BuildChangeSetRequestBody(changeSetIdString, "ADD", "", endpointsArray);
        //Build Header of the Batch Request
        var batchRequestBody = BuildChangeSetRequestHeader(boundryString, changeSetIdString, userPropertiesBatchBody);
        //Make the REST API call to the _api/$batch endpoint with the batch data
        console.log("==========================================================");
        console.log(userPropertiesBatchBody);
        console.log("==========================================================");
        console.log(batchRequestBody);
        console.log("==========================================================");
        var requestHeaders = {
            'X-RequestDigest': FormDigestValue,
            'Content-Type': "multipart/mixed; boundary=\"batch_" + boundryString + "\""
        };
        return fetch(rootUrl + "/_api/$batch", {
            method: "POST",
            headers: requestHeaders,
            credentials: "include",
            body: batchRequestBody
        }).then(function (r) { return r.text(); }).then(function (r) {
            //Convert the text response to an array containing JSON objects of the results
            var results = parseResponse(r);
            //Properties will be returned in the same sequence they were added to the batch request
            // for (var i = 0; i < endpointsArray.length; i++) {
            //     console.log(endpointsArray[i] + " is ", results[i]);
            // }
            return results;
        });
    };
    var internalBatch = function (_a) {
        //Reference from  Vardhman Despande blogs https://www.vrdmn.com/2016/06/sharepoint-online-get-userprofile.html
        var FormDigestValue = _a.FormDigestValue, rootUrl = _a.rootUrl, batchUrls = _a.batchUrls;
        //AccountName of the user
        //  var userAccountName = encodeURIComponent("i:0#.f|membership|user@yourtenant.onmicrosoft.com");
        //Collection of Endpoint url to fetch
        var endpointsArray = batchUrls; //["https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(212)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(213)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(214)"]
        //Unique identifier, will be used to delimit different parts of the request body
        var boundryString = getBoundryString();
        //Build Body of the Batch Request
        var userPropertiesBatchBody = buildBatchRequestBody(endpointsArray, boundryString);
        //Build Header of the Batch Request
        var batchRequestBody = buildBatchRequestHeader(userPropertiesBatchBody, boundryString);
        //Make the REST API call to the _api/$batch endpoint with the batch data
        var requestHeaders = {
            'X-RequestDigest': FormDigestValue,
            'Content-Type': "multipart/mixed; boundary=\"batch_" + boundryString + "\""
        };
        return fetch(rootUrl + "/_api/$batch", {
            method: "POST",
            headers: requestHeaders,
            credentials: "include",
            body: batchRequestBody
        }).then(function (r) { return r.text(); }).then(function (r) {
            //Convert the text response to an array containing JSON objects of the results
            var results = parseResponse(r);
            //Properties will be returned in the same sequence they were added to the batch request
            // for (var i = 0; i < endpointsArray.length; i++) {
            //     console.log(endpointsArray[i] + " is ", results[i]);
            // }
            return results;
        });
    };
    return { GetBatchAll: makeBatchRequests, PostBatchAll: makePostBatchRequests };
})();
//Usage of batch GET Multiple Request
// Step 1: Prepare Array or request Url e.g arr
// Step 2: Pass rootUrl or SiteUrl to generate RequestDigest
// var arr=["https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(212)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(213)", "https://brgrp.sharepoint.com/_api/Lists/Getbytitle('PlaceHolderList')/items(214)"]
// var rootUrl= "https://brgrp.sharepoint.com"
// Step 3: Pass information as below and that's it it will return result
//BatchUtils.GetBatchAll({rootUrl:rootUrl,batchUrls:arr}) .then(r=>console.log(r))
