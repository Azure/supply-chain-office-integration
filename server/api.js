'use strict';

const express = require('express');
const router = express.Router();
const bodyParser = require('body-parser');
const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
const http = require('http');
const url = require('url');

const development = process.env.NODE_ENV !== 'production';
const iberaServicesEndpoint = "https://localhost:443";
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";


router.use(bodyParser.json());

function handleRequest(xhr, body, callback) {
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if  (xhr.status === 200) {
        callback(JSON.parse(xhr.responseText));
      } else {
        callback({
          error: 'Request status: ' + xhr.status
        });
      }
    }
  };
  xhr.send(body && JSON.stringify(body) || null);
}
router.post('/proof', function (req, res) {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('POST', iberaServicesEndpoint + "/api/proof");
        xhr.setRequestHeader('Content-Type', 'application/json');
        handleRequest(xhr, req.body, function(result){
             res.send({result: result, error: result.error});
        });
    }
    catch (ex){
        res.status(503).send({error: ex.message});
    }
});
router.get('/proof', function (req, res) {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', iberaServicesEndpoint + "/api/proof?decrypt=true&tracking_id=" + req.query.tracking_id);
        xhr.setRequestHeader('Content-Type', 'application/json');
        handleRequest(xhr, {}, function(result){
            res.send({result: result, error: result.error});
        });
    }
    catch (ex){
        res.status(503).send({error: ex.message});
    }
});
router.get('/key', function (req, res) {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', iberaServicesEndpoint + "/api/key?key_id=" + req.query.key_id);
        xhr.setRequestHeader('Content-Type', 'application/json');
        handleRequest(xhr, {}, function(result){
            res.send({result: result, error: result.error});
        });
    }
    catch (ex){
        res.status(503).send({error: ex.message});
    }
});
router.get('/hash', function (req, res) {
    http.get(url.parse(req.query.url), function(res) {
        var data = [];

        res.on('data', function(chunk) {
            data.push(chunk);
        }).on('end', function() {
            //at this point data is an array of Buffers
            //so Buffer.concat() can make us a new Buffer
            //of all of them together
            var buffer = Buffer.concat(data);
            res.send({result: sha256(buffer).ToUpperCase(), error: result.error});
        }).on('error', function() {
            res.status(503).send({error: "error hashing"});
        })
    });
});
module.exports = router;
