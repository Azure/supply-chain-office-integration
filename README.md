## Running in localhost

Before running for the first time, generate a certificate and a key with the script borrowed from [here](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth/blob/7d125dd2862c629ee10baddffe981e84f0ed3b2d/ss_certgen.sh).

On Linux, Mac and Git Bash for Windows

```
$ bash ss_certgen.sh
```
On Cygwin for Windows

```
$ bash -o igncr ss_certgen.sh
```

To install dependencies:

```
$ npm install
```

To start the server:

```
$ npm start
```

Open Web browser https://localhost:8443/ and make the certificate trusted by adding it to the trusted root authorities.

To test that the REST API is accessible and working, you can issue a request to the ping endpoint and expect a hash to be returned:

``` 
$ curl https://localhost:8443/api/key
```


## Configuration 

To run locally, copy the file `dev.sample.json` in the `config` folder and create a new file called `dev.private.json`.

Fill in the following values:

`IBERA_SERVICES_ENDPOINT`: The URL for the iBera services service - the endpoint needs to support SSL
`DOCUMENT_SERVICES_ENDPOINT`: The URL for the iBera document service - the endpoint needs to support SSL 
When deploying to production, fill these values in the environment variables.

Also in `static/client.env.js`, set `documentServiceUrl` to your local document service endpoint.


## Deployment

Run the gulp dist task and provide the URL behind which you are deploying, for example:

```
$ gulp dist --url https://ibera-web-app.azurewebsites.net/
```

Above command rewrites the manifest to point to the correct resources.

The result is a dist folder that you can push to your hosting environment, for example:

```
$ cd dist
# Below in case you want to build the dependencies locally
$ npm install --production
$ git init .
$ git add *
$ git commit -m "Deployment"
$ git push https://ibera-web-app.scm.azurewebsites.net:443/ibera-web-app.git master
```
