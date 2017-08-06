'use strict';

// TODO: asb Beat- what's that?
process.env.STAMPERY_TOKEN = "65002b66-089a-4aad-eb00-962c5c4e0e59";


var util = require('util');
var express = require('express');
var http = require('http');
var https = require('https');
var fs = require('fs');
var path = require('path');
var bodyParser = require('body-parser');
var expressValidator = require('express-validator');

var api = require('./api');

var port = process.env.PORT || 8443;
var development = process.env.NODE_ENV !== 'production';
var app = express();

var serverOptions = {};

if (development) {
  serverOptions.cert = fs.readFileSync('./cert/server.crt');
  serverOptions.key = fs.readFileSync('./cert/server.key');
}

app.use(bodyParser.json());
app.use(expressValidator());

// middleware to log all incoming requests
app.use((req, res, next) => {
	console.log(`url: ${req.method} ${req.originalUrl} ${util.inspect(req.body || {})}`);
	return next();
});

// attach API to server
app.use('/api', api);

app.use('/', express.static(path.join(__dirname, 'static')));

https.createServer(serverOptions, app).listen(port, err => {
	if (err) return console.error(err);
	console.info(`server is listening on port ${port}`);
});
