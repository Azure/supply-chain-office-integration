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
const uuidv4 = require('uuid/v4');

var utils = require('../utils');
var auth = require('../auth.js');
var app = express();

const iberaServicesEndpoint = config.IBERA_SERVICES_ENDPOINT;
const azureStorageConnectionString = config.STORAGE_CONNECTION_STRING;

const USER_ACCOUNT_HEADER_KEY = 'user-account';
const CONTAINER_NAME = 'attachments';

var azureBlobService = azureStorage.createBlobService(azureStorageConnectionString);

app.get('/config', async (req, res) => {
  try {
    var result = {documentServiceUrl: config.DOCUMENT_SERVICES_ENDPOINT}

    console.log(`sending configuration: ${util.inspect(result)}`);
    return res.json(result);
  } catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
  }
});


app.use(async (req, res, next) => {
  try{
    // Verify User Account header:
    var userAccount = req.headers[USER_ACCOUNT_HEADER_KEY];
    if (!userAccount) return res.status(HttpStatus.BAD_REQUEST).json({error: USER_ACCOUNT_HEADER_KEY + ` request header is missing`});

    // Authentication request JWT:
    var authHeader = req.headers['authorization'] || '';
    var elements = authHeader.split(' ');
    if (elements.length !== 2) return res.status(HttpStatus.BAD_REQUEST).json({error: `invalid authorization token`});
    
    var requestJwt = elements[1];
    var jwtVerified = await auth.verifyJwt(requestJwt);
    if(!jwtVerified) return res.status(HttpStatus.UNAUTHORIZED).json({error: `Request is unauthorised with this authorization token`});
    
    console.log(`request authenticated.`);

    var userExchangeId = auth.parseAppctx(requestJwt).msexchuid;
    req.userInfo = {
      account: userAccount, 
      exchangeId: userExchangeId
    };
    
    return next();

  } catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
  }
});


app.post('/attachment', async(req, res) => {
  if(!validate(req.body, schema.attachment.post).valid){
    return res.status(HttpStatus.BAD_REQUEST).json({ error: `invalid schema - expected schema is ${util.inspect(schema.attachment.put)}` });
  }

  try {
    var userId = req.userInfo.account;
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
      var blobName = uuidv4() + "/" + encodeURIComponent(fileName);

      var binaryStream = intoStream(binaryData);

      if (req.body.upload) {
        await utils.callAsyncFunc(azureBlobService, 'createBlockBlobFromStream', CONTAINER_NAME, blobName, binaryStream, binaryData.length);
      }

      attachmentProcessingDetails.push({
        name: fileName,
        hash: contentHash,
        blobName: blobName
      });
    }

    return res.json({attachmentProcessingDetails: attachmentProcessingDetails});
    
  } catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({error: err.message});
  }

});

app.put('/proof', async(req, res) => {
  try {
    var uri = iberaServicesEndpoint + `/api/proof`;
    var results = [];
    if (req.body.length > 0){
      for (var i = 0; i < req.body.length; i++){
        req.body[i].userId = req.userInfo.account; //Add user id to the proof
        var result = await request({
          method: 'PUT',
          uri,
          body: req.body[i],
          json: true
        });
        results.push(result);
      }
    }
    
    console.log(`results: ${util.inspect(results)}`);
    return res.json(results);
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
    
    // trackingId is encoded. leave it encoded since we also use it as part of the URL in the request
    var trackingId = req.params.trackingId;
    if (decodeURIComponent(trackingId) === trackingId) {
      trackingId = encodeURIComponent(trackingId);
    }

    var decrypt = req.sanitizeQuery('decrypt').toBoolean();

    var path = iberaServicesEndpoint + `/api/proof/${trackingId}?decrypt=${decrypt}`;

    try {
      var result = await request.get(path, {json: true, headers: {'user-id': req.userInfo.account}});
      await addSasUrls(result);
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

async function addSasUrls(getProofsResponse){
  for (var i=0; i<getProofsResponse.proofs.length; i++){
    var blobName = getProofsResponse.proofs[i].encryptedProof.blobName;
    var sasToken = getSAS(CONTAINER_NAME, azureBlobService, blobName);
    var sasUrl = azureBlobService.getUrl(CONTAINER_NAME, blobName, sasToken, true);

    getProofsResponse.proofs[i].sasUrl = sasUrl;
  }
}

function getSAS(CONTAINER_NAME, blobSvc, blobName) {
  var sharedAccessPolicy = {
    AccessPolicy: {
      Start: azureStorage.date.minutesFromNow(-1),
      Expiry: azureStorage.date.minutesFromNow(2),
      Permissions: azureStorage.BlobUtilities.SharedAccessPermissions.READ
    }
  };
  var sasToken = blobSvc.generateSharedAccessSignature(CONTAINER_NAME, blobName, sharedAccessPolicy);
  console.log('sasToken', sasToken);
  return sasToken;
}


app.get('/key/:keyId', async (req, res) => {
  try {
    req.checkParams('keyId', 'Invalid keyId').notEmpty();
    var errors = await req.getValidationResult();
    if (!errors.isEmpty()) {
      return res.status(HttpStatus.BAD_REQUEST).json({error: `there have been validation errors: ${util.inspect(errors.array())}`});
    }

    // keyId is encoded. leave it encoded since we also use it as part of the URL in the request
    var keyId = req.params.keyId;
    if (decodeURIComponent(keyId) === keyId) {
      keyId = encodeURIComponent(keyId);
    }

    var userId = req.userInfo.account;

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