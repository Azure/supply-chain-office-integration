/*
 * Please refer to the dev.sample.json file.
 * Copy this file and create a new file named "dev.private.json".
*/

var nconf = require('nconf');
var path = require('path');

var envFile = '';
if (process.env.NODE_ENV == 'production') { 
  envFile = path.join(__dirname, 'prod.private.json');
} else {
  envFile = path.join(__dirname, 'dev.private.json');
}
console.log(`using env file: ${envFile}`);

var nconfig = nconf.env().file({ file: envFile });


// centric place to read and parse all configuration values
var config = {};

// validate configuration
var params = [
  'SUPPLY_CHAIN_SERVICES_ENDPOINT',
  'OUTLOOK_SERVICE_ENDPOINT',
  'STORAGE_CONNECTION_STRING'
];

// read configuration values from nconfig
params.forEach(param => config[param] = nconfig.get(param));

// validate confiuration
params.forEach(param => validate(param));

function validate(param) {
  if (!config[param]) {
    console.error(`EXISTING PROCESS: configuration param missing: '${param}'`);
    process.nextTick(() => process.exit(1));
  }
}


config.nconfig = nconf;

module.exports = config;