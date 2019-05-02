class SPRest {
  rootUrl: any;
  constructor(rootWeb) {
    this.rootUrl = rootWeb;
  }

  Utils = (() => {
    /**
     * Reference or motivation link : https://github.com/omkarkhair/sp-jello/blob/master/lib/jello.js
     * https://github.com/abhishekseth054/SharePoint-Rest
     * https://github.com/gunjandatta/sprest
     */
    var reqUrl =
      this.rootUrl + "/_api/web/lists/getbytitle('AWSResponseList')/items";
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

    console.log(this.rootUrl);
    let isOdataVerbose = true;
    let _rootUrl = this.rootUrl;
    //Here "Accept" and "Content-Type" header must require for Sharepoint Onlin REST API
    let _payloadOptions = {
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
    const getItemTypeForListName = name => {
      return (
        "SP.Data." +
        name.charAt(0).toUpperCase() +
        name
          .split(" ")
          .join("")
          .slice(1) +
        "ListItem"
      );
    };
    const Get = url => {
      var _localPayload = _payloadOptions;
      _localPayload.method = "GET";
      //Internally if body is set for GET request then need to remove it by setting undefined
      // otherwise it return with error : Failed to execute 'fetch' on 'Window': Request with GET/HEAD method cannot have body
      _localPayload.body = undefined;
      delete _localPayload.headers["IF-MATCH"];
      delete _localPayload.headers["X-HTTP-Method"];
      console.log(_localPayload);
      return fetch(url, _localPayload).then(r => r.json());
    };

    const getRequestDigest = (url = "") => {
      if (url) {
        url = this.rootUrl;
      }
      url += "/_api/contextinfo";

      var _localPayload = _payloadOptions;
      _localPayload.method = "POST";
      return fetch(url, _payloadOptions).then(r => r.json());
    };
    const postWithRequestDigestExtension = (url, { headers = {}, payload }) => {
      return getRequestDigest().then(token => {
        // payload.requestDigest = token.d.GetContextWebInformation.FormDigestValue;
        headers["X-RequestDigest"] =
          token.d.GetContextWebInformation.FormDigestValue;
        return PostExtension(url, { headers, payload });
      });
    };
    const postWithRequestDigest = (url, payload) => {
      return getRequestDigest().then(token => {
        payload.requestDigest =
          token.d.GetContextWebInformation.FormDigestValue;
        return Post(url, payload);
      });
    };
    //Need to refactor below method and need to merge in single method postWithRequestDigestExtension
    //Based on Action Type in headers using switch case add or remove extra headers
    // Try to add header name same as standard name so it can be replaced with for loop e.g. _extraHeaders
    const updateWithRequestDigest = (url, payload) => {
      return getRequestDigest(payload.rootUrl + "/_api/contextinfo").then(
        token => {
          payload.requestDigest =
            token.d.GetContextWebInformation.FormDigestValue;
          payload._extraHeaders = {
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
          };
          //For Update operation or Merge Operation no response will return only status will return for http request
          payload.isNoJsonResponse = true;
          return Post(url, payload);
        }
      );
    };
    const deleteWithRequestDigest = (url, payload) => {
      return getRequestDigest(payload.rootUrl + "/_api/contextinfo").then(
        token => {
          payload.requestDigest =
            token.d.GetContextWebInformation.FormDigestValue;
          payload._extraHeaders = {
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
          };
          //For Update operation or Merge Operation no response will return only status will return for http request
          payload.isNoJsonResponse = true;
          return Post(url, payload);
        }
      );
    };

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
    const PostExtension = (url, { headers, payload }) => {
      let _localPayload: any = {};
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
        return fetch(url, _localPayload).then(r => r);
      } else {
        return fetch(url, _localPayload).then(r => r.json());
      }
    };
    /**
     * Generate GUID in javascript
     * Reference from : https://github.com/andrewconnell/sp-o365-rest/blob/master/SpRestBatchSample/Scripts/App.js
     */
    function generateUUID() {
      var d = new Date().getTime();
      var uuid = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(
        /[xy]/g,
        function (c) {
          var r = (d + Math.random() * 16) % 16 | 0;
          d = Math.floor(d / 16);
          return (c == "x" ? r : (r & 0x7) | 0x8).toString(16);
        }
      );
      return uuid;
    }
    /**
 * Prepare batch request in Sharepoint Online
 * Reference from : https://github.com/andrewconnell/sp-o365-rest/blob/master/SpRestBatchSample/Scripts/App.js
 */
    const GenerateBatchRequest = ({
      listName = "",
      data,
      DigestValue = "",
      action = "ADD",
      itemIds = []
    }) => {
      // generate a batch boundary
      var batchGuid = generateUUID();
      // creating the body
      var batchContents = new Array();
      var changeSetId = generateUUID();
      // get current host
      // var temp = document.createElement('a');
      // temp.href = this.rootUrl;
      // var host = temp.hostname;
      // iterate through each employee
      ////TODO NEED TO SEPARATE BATCH OPERATION CREATION FOR ADD,UPDATE and DELETE OPERATION
      // MIXED INSERT and UPDATE Operation
      // MIXED ISERT ,UPDATE and DELETE Operation
      for (var _index = 0; _index < data.length; _index++) {
        //TODO for payload or "data"  need to generate or extract metadata based on type and batch action
        var _item = data[_index];
        if (action === "UPDATE") {
          _item.Title = "##Updated_" + _item.Title;
        }
        //Generate and prepare request url for each item
        switch (action) {
          case "ADD":
            var endpoint = generateRestRequest({ listName, type: "ListItem" });
            break;
          case "UPDATE":
            var endpoint = generateRestRequest({
              listName,
              type: "ListItem",
              Id: itemIds[_index]
            });
            _item.Title += _index;
            break;
          case "DELETE":
            var endpoint = generateRestRequest({
              listName,
              type: "ListItem",
              Id: itemIds[_index]
            });
            break;
        }
        // create the request endpoint

        // create the changeset
        batchContents.push("--changeset_" + changeSetId);
        batchContents.push("Content-Type: application/http");
        batchContents.push("Content-Transfer-Encoding: binary");
        batchContents.push("");
        if (action === "UPDATE") {
          batchContents.push("PATCH " + endpoint + " HTTP/1.1");
          batchContents.push("If-Match: *");
          batchContents.push("Content-Type: application/json;odata=verbose");
          batchContents.push("");
          batchContents.push(JSON.stringify(_item));
        } else if (action === "ADD") {
          batchContents.push("POST " + endpoint + " HTTP/1.1");
          batchContents.push("Content-Type: application/json;odata=verbose");
          batchContents.push("");
          batchContents.push(JSON.stringify(_item));
        } else if (action === "DELETE") {
          batchContents.push("DELETE " + endpoint + " HTTP/1.1");
          batchContents.push("If-Match: *");
        }
        //Commented POST request line and added code for UPDATE as well
        // batchContents.push("POST " + endpoint + " HTTP/1.1");
        // batchContents.push("Content-Type: application/json;odata=verbose");
        // batchContents.push("");
        // batchContents.push(JSON.stringify(_item));
        batchContents.push("");
      }
      // END changeset to create data
      batchContents.push("--changeset_" + changeSetId + "--");

      // batch body
      var batchBody = batchContents.join("\r\n");

      batchContents = [];

      // create batch for creating items
      batchContents.push("--batch_" + batchGuid);
      batchContents.push(
        'Content-Type: multipart/mixed; boundary="changeset_' +
        changeSetId +
        '"'
      );
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

      batchBody = batchContents.join("\r\n");

      // create the request endpoint
      var batchEndpoint = this.rootUrl + "/_api/$batch";

      // var batchRequestHeader = {
      //     'X-RequestDigest': DigestValue,
      //     'Content-Type': 'multipart/mixed; boundary="batch_' + batchGuid + '"'
      // };
      return getRequestDigest().then(r => {
        //SPECIAL CASE IS BATCH REQUEST MODE
        var batchRequestHeader = {
          "X-RequestDigest": r.d.GetContextWebInformation.FormDigestValue,
          "Content-Type": 'multipart/mixed; boundary="batch_' + batchGuid + '"'
        };
        return fetch(batchEndpoint, {
          method: "POST",
          headers: batchRequestHeader,
          body: batchBody
        }).then(r => r.text());
      });

      //return PostExtension(batchEndpoint,{headers:batchRequestHeader,payload:batchBody})
      //fetch(batchEndpoint,{method:'POST',headers:batchRequestHeader,body:batchBody}).then(r=>r.text()).then(r=>console.log(r))
      //OLD CODE Commented
      // create request
      // jQuery.ajax({
      //     url: endpoint,
      //     type: 'POST',
      //     headers: batchRequestHeader,
      //     data: batchBody,
      //     success: function(response) {

      //         var responseInLines = response.split('\n');

      //         //  $("#tHead").append("<tr><th>First Name</th><th>Last Name</th><th>Technology</th></tr>");

      //         for (var currentLine = 0; currentLine < responseInLines.length; currentLine++) {
      //             try {

      //                 var tryParseJson = JSON.parse(responseInLines[currentLine]);

      //                 console.log(tryParseJson);
      //             } catch (e) {
      //                 console.log("Error")
      //             }
      //         }
      //     },
      //     fail: function(error) {

      //     }
      // });
    };

    const generateRestRequest = ({ listName = "", Id = "", type, oDataOption = "", url = "" }) => {
      if (url) {
        return url;
      }
      let prepareRequest = `${this.rootUrl}/_api/web/lists`;
      switch (type) {
        case "ListItem":
          prepareRequest += `/getbytitle('${listName}')/items`;
          if (Id) {
            prepareRequest += `(${Id})`;
          }
          if (oDataOption) {
            prepareRequest += `?${oDataOption}`;
          }
          break;
        case "List":
          if (listName) {
            prepareRequest += `/getbytitle('${listName}')`;
          }
          if (oDataOption) {
            prepareRequest += `?${oDataOption}`;
          }
          //TODO
          break;
        case "MSGraph":
          //Special Case
         
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
    const ListItem = {
      Add: ({ }: any) => { },
      Update: ({ }: any) => { },
      Delete: ({ }: any) => { },
      GetItemById: ({ }: any) => { },
      GetAllItem: ({ }: any) => { }
    };
    const List = {
      Add: ({ }: any) => { },
      Update: ({ }: any) => {
        console.log("Implementation is pending");
      },
      Delete: ({ }: any) => {
        console.log("Implementation is pending");
      },
      GetByTitle: ({ }: any) => {
        console.log("Implementation is pending");
      },
      GetAll: ({ }: any) => { }
    };

    List.GetAll = ({ oDataOption = "", url = "" }) => {
      return Get(generateRestRequest({ listName: "", type: "List", oDataOption, url }));
    };
    List.GetByTitle = ({ listName, oDataOption, url }) => {
      return Get(generateRestRequest({ listName: listName, type: "List", oDataOption, url }));
    };
    List.Add = ({ listName, data }) => {
      let _reqUrl = generateRestRequest({ listName, type: "List" });
      let _prePayload = preparePayloadData({
        action: "ADD",
        type: "List",
        listName,
        data
      });
      return postWithRequestDigestExtension(_reqUrl, _prePayload);
    };

    ListItem.GetAllItem = ({ listName, oDataOption = "", url = "" }) => {

      return Get(generateRestRequest({ listName, type: "ListItem", oDataOption, url }));
    };

    ListItem.GetItemById = ({ listName, Id, oDataOption = "", url = "" }) => {
      return Get(generateRestRequest({ listName, Id, type: "ListItem", oDataOption, url }));
    };
    // Add/Update/Delete require same types of metaData as well as payload only difference is Headers

    ListItem.Add = ({ listName, data }) => {
      let _reqUrl = generateRestRequest({ listName, type: "ListItem" });
      let _prePayload = preparePayloadData({
        action: "ADD",
        type: "ListItem",
        listName,
        data
      });
      return postWithRequestDigestExtension(_reqUrl, _prePayload);
    };

    ListItem.Update = ({ listName, Id, data }) => {
      var _reqUrl = generateRestRequest({ listName, Id, type: "ListItem" });
      let _prePayload = preparePayloadData({
        action: "UPDATE",
        type: "ListItem",
        listName,
        data
      });
      return postWithRequestDigestExtension(_reqUrl, _prePayload);
    };

    ListItem.Delete = ({ listName, Id, data }) => {
      let _reqUrl = generateRestRequest({ listName, Id, type: "ListItem" });
      let _prePayload = preparePayloadData({
        action: "DELETE",
        type: "ListItem",
        listName,
        data
      });
      return postWithRequestDigestExtension(_reqUrl, _prePayload);
    };

    let preparePayloadData = ({ listName, data, action, type }) => {
      let payload: any = {
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
        for (let _key in data) {
          payload.data[_key] = data[_key];
        }
      }
      //TODO Verify about extra header

      // //Pass body data as stringyfy;
      // // _payloadOptions.body = JSON.stringify(_body);
      // // _payloadOptions.body.Title = title;

      //    return postWithRequestDigest(_reqUrl);
      let _headers: any = {};
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

      return { headers: _headers, payload };
    };

    return {
      ListItem: ListItem,
      Get: Get,
      getRequestDigest: getRequestDigest,
      Post: postWithRequestDigest,
      Update: updateWithRequestDigest,
      Delete: deleteWithRequestDigest,
      BatchInsert: GenerateBatchRequest
    };
  })();
}
