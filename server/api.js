'use strict';

const express = require('express');
const router = express.Router();
const bodyParser = require('body-parser');
const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

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
  xhr.onerror = function () {
    callback({
      error: 'Request error'
    });
  };
  xhr.send(body && JSON.stringify(body) || null);
}
router.post('/proof', function (req, res) {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('POST', iberaServicesEndpoint + "/api/proof");
        xhr.setRequestHeader('Content-Type', 'application/json');
        handleRequest(xhr, {
            "tracking_id" : "tracking_id_10",
            "encrypted_proof" : "YmFzZTY0IGRlY29kZXI=",
            "public_proof" :  "{producer_id:farmer2, email:test@farmer2.de}"
        }, function(result){
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
module.exports = router;
