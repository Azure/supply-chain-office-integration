'use strict';

var util = require('util');
var express = require('express');
var HttpStatus = require('http-status-codes');
var validate = require('jsonschema').validate;
var schema = require('./schema');
var request = require('request-promise');
var config = require('../config');

const ews = require('ews-javascript-api');
const azureStorage = require('azure-storage');
const sha256 = require('sha256');
const intoStream = require('into-stream');

var utils = require('../utils');

var app = express();

const iberaServicesEndpoint = config.IBERA_SERVICES_ENDPOINT;
const azureStorageConnectionString = config.STORAGE_CONNECTION_STRING;

const USER_TOKEN_HEADER_KEY = 'user-token';
const CONTAINER_NAME = 'attachments';

var azureBlobService = azureStorage.createBlobService(azureStorageConnectionString);

//Will be implemented next,  https://amitu-ted.visualstudio.com/DefaultCollection/iBera/_workitems/edit/11
async function getUserId(userToken) {
  return 'demo-user-001';
}


app.get('/config', async (req, res) => {
  try {
    var result = {documentServiceUrl: config.DOCUMENT_SERVICES_ENDPOINT}

    console.log(`sending configuration: ${util.inspect(result)}`);
    return res.json(result);
  } catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
  }
});

app.post('/attachment', async(req, res) => {
  var userToken = req.headers[USER_TOKEN_HEADER_KEY];
  if (!userToken) {
    return res.status(HttpStatus.BAD_REQUEST).json({error: CONTAINER_NAME + ` request header is missing`});
  }

  if(!validate(req.body, schema.attachment.post).valid){
    return res.status(HttpStatus.BAD_REQUEST).json({ error: `invalid schema - expected schema is ${util.inspect(schema.attachment.put)}` });
  }

  var userId = await getUserId(userToken);
  try {
    var attachmentProcessingDetails = [];
    var exch = new ews.ExchangeService(ews.ExchangeVersion.Exchange2013);
    exch.Url = new ews.Uri(req.body.ewsUrl);
    exch.Credentials = new ews.OAuthCredentials(req.body.attachmentToken);

    var attachmentIds = req.body.attachments.map(attachment => attachment.id);

    // NOTE: Since the exch.GetAttachments() API returns the full attachment in a base64 format, 
    // and not as a stream, we assume that the attachments are small enough to fit into a memory 
    // stream before sending them to the blob storage.
    var getAttachmentRequest = await exch.GetAttachments(attachmentIds, ews.BodyType.Text, null);
    var attachments = getAttachmentRequest.responses;

    await utils.callAsyncFunc(azureBlobService, 'createContainerIfNotExists', CONTAINER_NAME);

    // Handle responses (for every attachemnt there is a response in reponses):
    for (var i = 0; i < attachments.length; i++) {
      var fileName = attachments[i].attachment.name;
      var base64Content = attachments[i].attachment.base64Content;
      var binaryData = Buffer.from(base64Content, 'base64');
      var contentHash = sha256(binaryData);
      var blobName = encodeURIComponent(userId) + "/" + encodeURIComponent(contentHash) + "/" + encodeURIComponent(fileName);

      var binaryStream = intoStream(binaryData);

      if (req.body.upload) {
        await utils.callAsyncFunc(azureBlobService, 'createBlockBlobFromStream', CONTAINER_NAME, blobName, binaryStream, binaryData.length);
        var sasToken = getSAS(CONTAINER_NAME, azureBlobService, {name: blobName});
        var sasUrl = azureBlobService.getUrl(CONTAINER_NAME, blobName, sasToken, true);
      }

      attachmentProcessingDetails.push({
        name: fileName,
        hash: contentHash,
        sasUrl: sasUrl
      });
    }

    return res.json({attachmentProcessingDetails: attachmentProcessingDetails});
    
  } catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
  }

});

function getSAS(CONTAINER_NAME, blobSvc, opts) {
  var sharedAccessPolicy = {
    AccessPolicy: {
      Start: azureStorage.date.minutesFromNow(-1),
      Expiry: azureStorage.date.minutesFromNow(2),
      Permissions: azureStorage.BlobUtilities.SharedAccessPermissions.READ
    }
  };
  var sasToken = blobSvc.generateSharedAccessSignature(CONTAINER_NAME, opts.name, sharedAccessPolicy);
  console.log('sasToken', sasToken);
  return sasToken;
}

app.put('/proof', async(req, res) => {
  try {
    if (!req.headers[USER_TOKEN_HEADER_KEY]) {
      return res.status(HttpStatus.BAD_REQUEST).json({error: CONTAINER_NAME + ` request header is missing`});
    }

    req.body.userId = await getUserId(req.headers[USER_TOKEN_HEADER_KEY]);

    var uri = iberaServicesEndpoint + `/api/proof`;
    var result = await request({
      method: 'PUT',
      uri,
      body: req.body,
      json: true
    });

    console.log(`got response: ${util.inspect(result)}`);
    return res.json(result);
  } catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
  }
});

app.get('/proof/:trackingId', async(req, res) => {
  try {
    req.checkParams('trackingId', 'Invalid trackingId').notEmpty();
    var errors = await req.getValidationResult();
    if (!errors.isEmpty()) {
      return res.status(HttpStatus.BAD_REQUEST).json({error: `there have been validation errors: ${util.inspect(errors.array())}`});
    }
    if (!req.headers[USER_TOKEN_HEADER_KEY]) {
      return res.status(HttpStatus.BAD_REQUEST).json({error: CONTAINER_NAME + ` request header is missing`});
    }

    var userId = await getUserId(req.headers[USER_TOKEN_HEADER_KEY]);

    // trackingId is encoded. leave it encoded since we also use it as part of the URL in the request
    var trackingId = req.params.trackingId;
    if (decodeURIComponent(trackingId) === trackingId) {
      trackingId = encodeURIComponent(trackingId);
    }

    var decrypt = req.sanitizeQuery('decrypt').toBoolean();

    var path = iberaServicesEndpoint + `/api/proof/${trackingId}?userId=${userId}&decrypt=${decrypt}`;

    try {
      var result = await request.get(path, {json: true});
    } catch (err) {
      if (err.statusCode === HttpStatus.NOT_FOUND) {
        // pass on the error we got from the services api
        return res.status(HttpStatus.NOT_FOUND).json(err.error);
      }

      throw err;
    }

    console.log(`got response: ${util.inspect(result)}`);
    res.json({result});
  } catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
  }

});


app.get('/key/:keyId', async (req, res) => {
  try {
    req.checkParams('keyId', 'Invalid keyId').notEmpty();
    var errors = await req.getValidationResult();
    if (!errors.isEmpty()) {
      return res.status(HttpStatus.BAD_REQUEST).json({error: `there have been validation errors: ${util.inspect(errors.array())}`});
    }

    if (!req.headers[USER_TOKEN_HEADER_KEY]) {
      return res.status(HttpStatus.BAD_REQUEST).json({error: CONTAINER_NAME + ` request header is missing`});
    }

    // keyId is encoded. leave it encoded since we also use it as part of the URL in the request
    var keyId = req.params.keyId;
    if (decodeURIComponent(keyId) === keyId) {
      keyId = encodeURIComponent(keyId);
    }

    var userId = await getUserId(req.headers[USER_TOKEN_HEADER_KEY]);

    var path = iberaServicesEndpoint + `/api/key/${keyId}?userId=${userId}`;
    var result = await request.get(path, {json: true});

    console.log(`got response: ${util.inspect(result)}`);
    return res.json(result);
  } catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
  }
});


// TODO: revisit code
app.get('/hash', (req, res) => {
  console.log(`in hash api- ${util.inspect(req)}`);
  var url = decodeURIComponent(req.query.url);
  console.log(`getting url: '${url}'`);
  return http.get(url.parse(req.query.url), res => {
    var data = [];
    return res
      .on('data', function(chunk) {
        data.push(chunk);
      })
      .on('end', function() {
        //at this point data is an array of Buffers
        //so Buffer.concat() can make us a new Buffer
        //of all of them together
        var buffer = Buffer.concat(data);
        res.send({
          result: sha256(buffer).ToUpperCase(),
          error: result.error
        });
      })
      .on('error', err => {
        return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
      });

  });
});

module.exports = app;