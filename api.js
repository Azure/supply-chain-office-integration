'use strict';

var util = require('util');
var express = require('express');
var HttpStatus = require('http-status-codes');
var validate = require('jsonschema').validate;
var request = require('request-promise');

var app = express();

const development = process.env.NODE_ENV !== 'production';
const iberaServicesEndpoint = "https://localhost:443";
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

app.put('/proof', async (req, res) => {
  
  try {

    // TODOL add validations to schema

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


app.get('/proof/:tracking_id', async (req, res) => {
  
  try {
    req.checkParams('tracking_id', 'Invalid tracking_id').notEmpty();
    var errors = await req.getValidationResult();
    if (!errors.isEmpty()) {
      return res.status(HttpStatus.BAD_REQUEST).json({ error: `there have been validation errors: ${util.inspect(errors.array())}` });
    }

    var trackingId = decodeURIComponent(req.params.tracking_id);
    var decrypt = req.sanitizeQuery('decrypt').toBoolean();

    var path = iberaServicesEndpoint + `/api/proof/${trackingId}?decrypt=${decrypt}`;
    var result = await request.get(path, { json: true });

    console.log(`got response: ${util.inspect(result)}`);
    res.json(result);
  }
  catch (err) {
    return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ error: err.message });
  }

});


app.get('/key/:key_id', async (req, res) => {
  try {

    req.checkParams('key_id', 'Invalid key_id').notEmpty();
    var errors = await req.getValidationResult();
    if (!errors.isEmpty()) {
      return res.status(HttpStatus.BAD_REQUEST).json({ error: `there have been validation errors: ${util.inspect(errors.array())}` });
    }

    var keyId = decodeURIComponent(req.params.key_id);

    var path = iberaServicesEndpoint + `/api/key/${keyId}`;
    var result = await request.get(path, { json: true });

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
