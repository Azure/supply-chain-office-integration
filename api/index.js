'use strict';

var util = require('util');
var express = require('express');
var HttpStatus = require('http-status-codes');
var validate = require('jsonschema').validate;
var request = require('request-promise');
var config = require('../config');

var app = express();

const development = process.env.NODE_ENV !== 'production';
const iberaServicesEndpoint = config.IBERA_SERVICES_ENDPOINT;
const documentServicesEndpoint = config.DOCUMENT_SERVICES_ENDPOINT;

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";



async function getUserId(userToken) {

  try {
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
    var result = await request.get(path, {
       json: true 
    });

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
