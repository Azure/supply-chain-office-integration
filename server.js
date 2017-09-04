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

var config = require('./config');
var utils = require('./utils');
var api = require('./api');

 // changing the port to anything other than 8443 -> please update 'static/schema.xml' file and also in the '/manifest.xml' http handler below
var port = process.env.PORT || 8443;

var app = express();
var serverOptions = {};

if (utils.isProd) {

	// in prod, enforce secured connections
	app.use((req, res, next) => {
		if (!req.headers['x-arr-ssl']) {
			return res.status(HttpStatus.BAD_REQUEST).json({ error: 'use https'});
		}
		return next();
	});
}
else { // dev

	// accept self-signed cdertificates only in development
	process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

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

// hijack request for manifest file- 
// render the URLs to point to the deployed domain
app.get('/manifest.xml', async (req, res) => {
	try {
		var content = (await utils.callAsyncFunc(fs, 'readFile', 'static/manifest.xml', 'utf-8')).result;

		// "https://localhost:8443" should always be used for development
		content = content.replace(new RegExp('https://localhost:8443', 'g'), config.OUTLOOK_SERVICE_ENDPOINT);
		
		res.header('content-type', 'application/xml');
		return res.end(content);
	}
	catch(err) {
		console.error(`error rendering manifest.xml file: ${err.message}`);
		return res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ error: err.message });
	}
});

app.use('/', express.static(path.join(__dirname, 'static')));

if (utils.isProd) {
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
