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
var isProd = process.env.NODE_ENV === 'production';
var app = express();

var serverOptions = {};


// middleware to log all incoming requests
app.use((req, res, next) => {
	console.log(`url: ${req.method} ${req.originalUrl} ${util.inspect(req.body || {})}, headers: ${util.inspect(req.headers)}`);
	return next();
});

if (isProd) {
	app.use((req, res) => {

		var proto = req.connection.encrypted ? 'https' : 'http';
    proto = req.headers['x-forwarded-proto'] || proto;
    proto = proto.split(/\s*,\s*/)[0];

		if (proto === 'http') {
			return res.status(403).json({ error: 'use https' });
		}
		return next();
	});
}
else {
  serverOptions.cert = fs.readFileSync('./cert/server.crt');
  serverOptions.key = fs.readFileSync('./cert/server.key');
}

app.use(bodyParser.json());
app.use(expressValidator());


// attach API to server
app.use('/api', api);

app.get('/', (req, res) => {
	return res.end(`iBera Outlook Add-In Service in on...`);
});

app.use('/', express.static(path.join(__dirname, 'static')));

if (isProd) {
	// in prod we will use Azure's certificate to use ssl.
	// so no need to use https here with a custom certificate for now.
	// we enforce https connections in the first middleware when running in prod (see above).
	http.createServer(app).listen(port, err => {
		if (err) return console.error(err);
		console.info(`server is listening on port ${port}`);
	});
}
else {
	https.createServer(serverOptions, app).listen(port, err => {
		if (err) return console.error(err);
		console.info(`server is listening on port ${port}`);
	});
}
