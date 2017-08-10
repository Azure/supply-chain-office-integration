'use strict';

// TODO: asb Beat- what's that?
process.env.STAMPERY_TOKEN = "65002b66-089a-4aad-eb00-962c5c4e0e59";

var util = require('util');
var express = require('express');
var cors = require('express-cors');
var http = require('http');
var https = require('https');
var fs = require('fs');
var path = require('path');
var bodyParser = require('body-parser');
var expressValidator = require('express-validator');
var HttpStatus = require('http-status-codes');

var api = require('./api');

var port = process.env.PORT || 8443;
var isProd = process.env.NODE_ENV === 'production';
var app = express();
var serverOptions = {};

if (isProd) {

	// in prod, enforce secured connections
	app.use((req, res, next) => {
		if (!req.headers['x-arr-ssl']) {
			return res.status(HttpStatus.BAD_REQUEST).json({ error: 'use https'});
		}
		return next();
	});
}
else {
  serverOptions.cert = fs.readFileSync('./cert/server.crt');
  serverOptions.key = fs.readFileSync('./cert/server.key');
}

app.use(cors());
app.use(bodyParser.json());
app.use(expressValidator());

// middleware to log all incoming requests
app.use((req, res, next) => {
	console.log(`url: ${req.method} ${req.originalUrl} ${util.inspect(req.body || {})}, headers: ${util.inspect(req.headers)}`);
	return next();
});

// attach API to server
app.use('/api', api);

app.get('/', (req, res) => {
	return res.end(`iBera Outlook Add-In Service in on...`);
});

app.use('/', express.static(path.join(__dirname, 'static')));

if (isProd) {
	// in prod we will use Azure's certificate to use ssl.
	// so no need to use https here with a custom certificate for now.
	// enforcing https in prod is being done on the first middleware (see above)
	http.createServer(app).listen(port, err => {
		if (err) return console.error(err);
		console.info(`server is listening on port ${port}`);
	});
}
else {
	// this is development environment, use a local ssl server with self signed certificates
	https.createServer(serverOptions, app).listen(port, err => {
		if (err) return console.error(err);
		console.info(`server is listening on port ${port}`);
	});
}

process.on('uncaughtException', err => {
	console.error(`uncaught exception: ${err.message}`);
	setTimeout(() =>  process.exit(1), 1000);
});
