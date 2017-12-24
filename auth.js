const util = require('util');
const request = require('request-promise');
const jsrsasign = require('jsrsasign');
const jsonwebtoken = require('jsonwebtoken');
const url = require('url');


async function verifyJwt(jwt) {
    try {
      var publicKey = await getPuclicKeyFromExchange(jwt);
      return jsrsasign.jws.JWS.verifyJWT(jwt,publicKey, {alg: ['RS256']});
    } catch(err) {
      var errMsg = `Error validating Json Web Token: ${util.inspect(err)}`;
      console.log(err);
      throw (new Error(errMsg));
    }
  }
  
  function parseAppctx(jwt) {
    var tokenDecoded = jsonwebtoken.decode(jwt);
    return JSON.parse(tokenDecoded.appctx);
  }
  
  function verifyAmurl(amurl) {
    var amurlObject = url.parse(amurl);
    var host = amurlObject.host;
    var protocol = amurlObject.protocol;
  
    if(!protocol.startsWith('https') || !(host.includes('microsoft.com') || host.includes('office365.com'))) {
      var errMsg = 'Unauthorized source of public key: ' + amurl + '. It must be https and with microsoft.com or office365.com in its hostname';
      console.log(errMsg);
      throw (new Error(errMsg));
    }
  }
  
  async function getPuclicKeyFromExchange(jwt) {
    try {
      // Get the URI where the public key exists:
      var amurl = parseAppctx(jwt).amurl;
    
      verifyAmurl(amurl);
      var exchangeResponse = await request.get(amurl);
    
      // Extract the Public Key from the response:
      if(exchangeResponse) {
        var exchangeResponseParsed = JSON.parse(exchangeResponse);
        if(exchangeResponseParsed.keys && exchangeResponseParsed.keys.length > 0) {
          // Search for the X509 Certificate key:
          var x509Keys = exchangeResponseParsed.keys.filter(key => key.keyvalue.type == 'x509Certificate');
          var publicKey = x509Keys[0].keyvalue.value;
        
          //Add header and footer to the key:
          return '-----BEGIN CERTIFICATE-----' + publicKey + '-----END CERTIFICATE-----';
        }
        
      }else {
        throw (new Error('Unable to get public key from this URL: ' + amurl));
      }
    } catch(err) {
      console.log(`Error getting public key: ${util.inspect(err)}`);
      throw (new Error(errMsg));
    }
  }

  var api = {
    verifyJwt,
    parseAppctx
  };
  
  module.exports = api;