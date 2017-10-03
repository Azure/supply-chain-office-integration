'use strict';

var util = require('util');
var express = require('express');
var HttpStatus = require('http-status-codes');
var validate = require('jsonschema').validate;
var request = require('request-promise');
var config = require('../config');

var ews = require('ews-javascript-api');
var azureStorage = require('azure-storage');
var sha256 = require('sha256');
const intoStream = require('into-stream');
const btoa = require('btoa');
const atob = require('atob');

var utils = require('../utils');

var app = express();

const iberaServicesEndpoint = config.IBERA_SERVICES_ENDPOINT;
//const documentServicesEndpoint = config.DOCUMENT_SERVICES_ENDPOINT;
const azureStorageConnectionString = 'DefaultEndpointsProtocol=https;AccountName=iberat2keys;AccountKey=VQeSopXnbY4qEW4l1oSzkYdRvyyTY5jxHE2yLPQ1BGldexp9lsUjmfqt39c0Wuq+lnNw7XYDJG4MkCtCfSeoVQ==;EndpointSuffix=core.windows.net';
//const azureStorageConnectionString = config.AZURE_OI_STORAGE_CONNECTION_STRING; //TODO: Add this connection string to ARM tempalate + modify the OI APP settings via the automation script

async function getUserId(userToken) {
  return 'demo-user-001';
 /* try {
    var uri = documentServicesEndpoint + `/api/user`;
    var result = await request({
      method: 'GET',
      uri,
      headers: { 'User-Token': userToken },
      json: true
    });
    console.log(`got response: ${util.inspect(result)}`);
    return result;
  }
  catch (err) {
    var errorMessage = `Could not retrieve userId from user token: ${userToken}`;
    console.log(errorMessage);
    throw new Error(errorMessage);
  }
  */
}


app.get('/config', async (req, res) => {
  
  try {
    var result = {
      documentServiceUrl: config.DOCUMENT_SERVICES_ENDPOINT
    }

    console.log(`sending configuration: ${util.inspect(result)}`);
    res.json(result);
  }
  catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ error: err.message });
  }
});

///////////////////////////////////////
app.post('/attachment', async (req, res) => {
  if (!req.headers['user-token']) {
    return res.status(HttpStatus.BAD_REQUEST).json({ error: `user-token request header is missing` });
  }

  var userName = 'demo-user-001'; //TODO: Get user according to the user token
  
  try{
    var attachmentProcessingDetails = [];
    var exch = new ews.ExchangeService(ews.ExchangeVersion.Exchange2013);
    exch.Url = new ews.Uri(req.body.ewsUrl);
    exch.Credentials = new ews.OAuthCredentials(req.body.attachmentToken);
    
    var attachmentIds = req.body.attachments.map(a => a.id);
    
    // TODO: Add comment about large files 
    var response = await exch.GetAttachments(attachmentIds,ews.BodyType.Text,null);
    var azureBlobService = azureStorage.createBlobService(azureStorageConnectionString);

    await utils.callAsyncFunc(azureBlobService,'createContainerIfNotExists','attachments');

    // Handle responses (for every attachemnt there is a response into reponses:
    for(var i=0; i<response.responses.length; i++){
      // TODO: Check errors in response
      var fileName = response.responses[i].attachment.name;
      var base64Content = response.responses[i].attachment.base64Content;
      

      var binaryData  =   new Buffer(base64Content, 'base64');
      var contentHash = sha256(binaryData);
      var blobName = userName + "/" + encodeURIComponent(contentHash) + "/" + fileName;
      
      var binaryStream = intoStream(binaryData);

      if(req.body.upload){
        // TODO: identify the file type (from Ami)
        await utils.callAsyncFunc(azureBlobService,'createBlockBlobFromStream','attachemnts',blobName,binaryStream,binaryData.byteLength);
        var sasToken = azureBlobService.generateSharedAccessSignature("attachemnts",blobName,{AccessPolicy:{Expiry:azureStorage.date.daysFromNow(7)}});
        var sasUrl = azureBlobService.getUrl("attachemnts",blobName,sasToken, true);
      }

      attachmentProcessingDetails.push(
        {
          name: fileName,
          hash: contentHash,
          sasUrl: sasUrl
        }
      );
    }

    res.json({
      attachemntsProcessed: attachmentProcessingDetails.length,
      attachmentProcessingDetails: attachmentProcessingDetails
    });
  }

  catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ error: err.message });
  }
    
});

///////////////////////////////////////

app.put('/proof', async (req, res) => {
  
  try {
    if (!req.headers['user-token']) {
        return res.status(HttpStatus.BAD_REQUEST).json({ error: `user-token request header is missing` });
    }

    req.body.userId = await getUserId(req.headers['user-token']);

    var uri = iberaServicesEndpoint + `/api/proof`;
    var result = await request({
      method: 'PUT',
      uri,
      body: req.body,
      json: true
    });

    console.log(`got response: ${util.inspect(result)}`);
    res.json(result);
  }
  catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ error: err.message });
  }
});

app.get('/proof/:trackingId', async (req, res) => {
  
  try {
    req.checkParams('trackingId', 'Invalid trackingId').notEmpty();
    var errors = await req.getValidationResult();
    if (!errors.isEmpty()) {
      return res.status(HttpStatus.BAD_REQUEST).json({ error: `there have been validation errors: ${util.inspect(errors.array())}` });
    }
    if (!req.headers['user-token']){
        return res.status(HttpStatus.BAD_REQUEST).json({ error: `user-token request header is missing` });
    }

    var userId = await getUserId(req.headers['user-token']);


    // trackingId is encoded. leave it encoded since we also use it as part of the URL in the request
    var trackingId = req.params.trackingId;
    if (decodeURIComponent(trackingId) === trackingId) {
      trackingId = encodeURIComponent(trackingId);
    }

    var decrypt = req.sanitizeQuery('decrypt').toBoolean();

    var path = iberaServicesEndpoint + `/api/proof/${trackingId}?userId=${userId}&decrypt=${decrypt}`;
    
    try {
      var result = await request.get(path, {
        json: true 
      });
    }
    catch (err) {
      if (err.statusCode === HttpStatus.NOT_FOUND) {
        // pass on the error we got from the services api
        return res.status(HttpStatus.NOT_FOUND).json(err.error);        
      }
      
      throw err;
    }

    console.log(`got response: ${util.inspect(result)}`);
    res.json({ result });
  }
  catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ error: err.message });
  }

});


app.get('/key/:keyId', async (req, res) => {
  try {

    req.checkParams('keyId', 'Invalid keyId').notEmpty();
    var errors = await req.getValidationResult();
    if (!errors.isEmpty()) {
      return res.status(HttpStatus.BAD_REQUEST).json({ error: `there have been validation errors: ${util.inspect(errors.array())}` });
    }

    if (!req.headers['user-token']){
        return res.status(HttpStatus.BAD_REQUEST).json({ error: `user-token request header is missing` });
    }
    
    // keyId is encoded. leave it encoded since we also use it as part of the URL in the request
    var keyId = req.params.keyId;
     if (decodeURIComponent(keyId) === keyId) {
      keyId = encodeURIComponent(keyId);
    }

    var userId =await getUserId(req.headers['user-token']);

    var path = iberaServicesEndpoint + `/api/key/${keyId}?userId=${userId}`;
    var result = await request.get(path, {
       json: true 
    });

    console.log(`got response: ${util.inspect(result)}`);
    res.json(result);
  }
  catch (ex) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ error: err.message });
  }
});


// TODO: revisit code
app.get('/hash', (req, res) => {
  console.log(`in hash api- ${util.inspect(req)}`);

  var url = decodeURIComponent(req.query.url);
  console.log(`getting url: '${url}'`);

  return http.get(url.parse(req.query.url), res => {
    var data = [];

    return res.on('data', function(chunk) {
      data.push(chunk);
    })
    .on('end', function() {
      //at this point data is an array of Buffers
      //so Buffer.concat() can make us a new Buffer
      //of all of them together
      var buffer = Buffer.concat(data);
      res.send({result: sha256(buffer).ToUpperCase(), error: result.error});
    })
    .on('error', err => {
      return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ error: err.message });
    });

  });
});


module.exports = app;
