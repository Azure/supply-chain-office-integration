// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
// See full license at the bottom of this file.


// The initialize function is required for all add-ins.
Office.initialize = function () {
  jQuery(document).ready(function () {
    console.log('JQuery initialized');
  });
};

console.log('loading supply-chain add-in');

// TODO move to configuration retrieved from the server
const containerName = "attachments";

const beginProofString = "-----BEGIN PROOF-----";
const endProofString = "-----END PROOF-----";


function httpRequest(opts, cb) {

  // get user token, add to headers and invoke the http request 
  return getUserIdentityToken(function (err, token) {
    if (err) return cb(err);

    opts.headers = opts.headers || {};
    if (!opts.headers['User-Token']) {
      opts.headers['User-Token'] = token;
    }

    console.log('calling', opts.method, opts.url, opts.data ? JSON.stringify(opts.data) : '');

    opts.success = function (data, textStatus) {
      console.log('got data:', data, textStatus);
      return cb(null, data);
    }

    opts.error = function (xhr, textStatus, errorThrown) {
      console.log('got error:', textStatus, errorThrown);
      var msg = 'error invoking http request';

      // override message if we got an error message from the server
      var response;
      try {
        response = JSON.parse(xhr.responseText);
      } catch (err) {
        console.warn('error parsing object: ', xhr.responseText);
      }

      if (response && response.error) {
        msg = response.error;
      }

      return cb(new Error(msg));
    }

    return $.ajax(opts);

  });
}


function putProof(proof, cb) {
  console.log('adding proof:', proof);

  return httpRequest({
    method: 'PUT',
    contentType: "application/json; charset=utf-8",
    url: '/api/proof',
    data: JSON.stringify(proof),
    dataType: 'json'
  }, cb);
}

function getKey(keyId, cb) {
  console.log('getting key for keyId', keyId);

  if (keyId === decodeURIComponent(keyId)) {
    keyId = encodeURIComponent(keyId);
  }

  return httpRequest({
    method: 'GET',
    url: '/api/key/' + keyId
  }, cb);
}

function getProof(trackingId, cb) {
  console.log('getting proof for trackingId', trackingId);

  if (trackingId === decodeURIComponent(trackingId)) {
    trackingId = encodeURIComponent(trackingId);
  }

  return httpRequest({
    method: 'GET',
    url: '/api/proof/' + trackingId
  }, cb);
}

function getHash(url, cb) {
  console.log('getting hash for url', url);

  return httpRequest({
    method: 'GET',
    url: '/api/hash?url=' + encodeURIComponent(url)
  }, cb);
}

function getUserIdentityToken(cb) {
  return Office.context.mailbox.getUserIdentityTokenAsync(function (userToken) {
    if (userToken.error) return cb(userToken.error);
    return cb(null, userToken.value);
  });
}

function getClientConfiguration(cb) {
  console.log('getting configuration from server');

  return httpRequest({
    method: 'GET',
    url: '/api/config'
  }, cb);
}

function storeAttachments(event) {
  console.log('storeAttachments called');
  return processAttachments(true, function (err, response) {
    if (err) return showMessage("Error: " + err.message, event);
    console.log('got response', response);

    var trackingIds = [];
    if (response.attachmentProcessingDetails) {
      for (i = 0; i < response.attachmentProcessingDetails.length; i++) {

        var ad = response.attachmentProcessingDetails[i];
        var proof = {
          proofToEncrypt: {
            sasUrl: ad.sasUrl,
            documentName: ad.name
          },
          publicProof: {
            documentHash: ad.hash
          }
        };

        return putProof(proof, function (err, response) {
          if (err) return showMessage(err.message, event);

          trackingIds.push(response.trackingId);

          Office.context.mailbox.item.displayReplyForm(JSON.stringify(trackingIds));
          return showMessage("Attachments processed: " + JSON.stringify(trackingIds), event);
        });
      }
    }
  });
}

function processAttachments(isUpload, cb) {

  console.log('processing attachments, isUpload:', isUpload);

  if (!Office.context.mailbox.item.attachments) {
    return cb(new Error("Not supported: Attachments are not supported by your Exchange server."));
  }

  if (!Office.context.mailbox.item.attachments.length) {
    return cb(new Error("No attachments: There are no attachments on this item."));
  }

  return Office.context.mailbox.getCallbackTokenAsync(function (attachmentTokenResult) {
    console.log('getCallbackTokenAsync callback result:', attachmentTokenResult);
    if (attachmentTokenResult.error) return cb(attachmentTokenResult.error);

    return getClientConfiguration(function (err, config) {
      if (err) return cb(err);

      var data = {};
      data.ewsUrl = Office.context.mailbox.ewsUrl;
      data.attachments = [];
      data.containerName = containerName;
      data.upload = isUpload;
      data.attachmentToken = attachmentTokenResult.value;

      // extract attachment details 
      for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
        var attachment = Office.context.mailbox.item.attachments[i];
        attachment = attachment._data$p$0 || attachment.$0_0;

        if (attachment) {
          // I copied this line from the msdn example - not sure why first stringify and then parse the attachment
          // TODO: check this. probably the origin intention was to create a new object. but I don't see why we need this.
          data.attachments[i] = JSON.parse(JSON.stringify(attachment));
        }
      }

      // **************************************************************************************************
      // TODO: remove, this is a temporary bypassing the document service until Beat brings it online
      /*
      return cb(null, {
        attachmentProcessingDetails: [
          {
            url: 'http://...',
            sasToken: 'some token',
            name: 'some name',
            hash: 'the hash!'
          }
        ]
      });
      */
      // **************************************************************************************************


      return httpRequest({
        // url: config.documentServiceUrl + "/api/Attachment",
        url: "/api/attachment",
        method: 'POST',
        contentType: "application/json; charset=utf-8",
        data: JSON.stringify(data),
        dataType: 'json',
      }, function (err, response) {
        if (err) return cb(err);

        // in this case the document service might return a result that contains an error, so also need to check this specifically
        // TODO: revisit api on document service after rewriting in Node.js.
        // if there's an error it should send back a statusCode != 200 to indicate that
        if (response.isError) return cb(new Error('error uploading document: ' + response.message));

        return cb(null, response);
      });
    });
  });
}

function getFirstAttachmentHash(cb) {

  return processAttachments(true, function (err, response) {
    if (err) return cb(err);
    console.log('got response', response);

    if (!response.attachmentProcessingDetails || !response.attachmentProcessingDetails.length) {
      console.error('hash is not available');
      return cb(new Error('hash not available'));
    }

    var hash = response.attachmentProcessingDetails[0].hash;
    return cb(null, {
      hash: hash
    });

  });
}


// TODO: revisit&rewrite this function
function validateProof(event) {
  return Office.context.mailbox.item.body.getAsync('text', {}, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      return showMessage(result.error, event);
    }

    try {
      var body = result.value;
      if (body.search(beginProofString) === -1 || body.search(endProofString) === -1) {
        return showMessage("No proofs to validate found in email", event);
      }

      var proofsStep1Array = body.split(beginProofString);
      var proofsStep2Array = proofsStep1Array[1].split(endProofString);

      try {
        var proofs = JSON.parse(proofsStep2Array[0]);
      } catch (err) {
        console.error('invalid json', proofsStep2Array[0]);
        return showMessage("Invalid json", event);
      }

      //var proofs = proofsObj.proofs;

      if (!proofs.length) {
        return showMessage("no proofs found", event);
      }

      for (var i in proofs) {

        var trackingId = proofs[i].trackingId;
        return getProof(trackingId, function (err, result) {
          console.log('get proof from chain:', err, result);
          if (err) {
            return showMessage("error retrieving the proof from blockchain for validation - trackingId: " + trackingId + " error: " + err.message, event);
          }

          if (!result) {
            return showMessage("error retrieving the proof from blockchain for validation - trackingId: " + trackingId, event);
          }

          var proofFromChain = result.result.proofs[0];
          var proofToEncryptStr = JSON.stringify(proofs[0].encryptedProof);
          var hash = sha256(proofToEncryptStr);

          if (proofFromChain.publicProof.encryptedProofHash !== hash.toUpperCase()) {
            return showMessage("NOT valid proof for trackingId: " + trackingId, event);
          }

          if (!proofFromChain.publicProof.publicProof || !proofFromChain.publicProof.publicProof.documentHash) {
            return showMessage("Valid proof with NO attachment for trackingId: " + trackingId, event);
          }

          return getFirstAttachmentHash(function (err, result) {
            console.log('retrieving first attachment hash:', err, result);
            if (err) {
              return showMessage("error retrieving first attachment hash - trackingId: " + trackingId + " error: " + err.message, event);
            }

            var hash = result.hash;
            if (proofFromChain.publicProof.publicProof.documentHash === hash) {
              return showMessage("Valid proof with attachment for trackingId: " + trackingId, event);
            }

            return showMessage("Valid proof BUT attachment NOT valid for trackingId: " + trackingId, event);

          });
        });
      }
    } catch (ex) {
      return showMessage(ex.message, event);
    }
  });
}

function provideProof(event) {
  return Office.context.mailbox.item.body.getAsync('text', {}, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      return showMessage(result.error, event);
    }

    var guids = extractGuidsFromText(result.value);
    if (!guids.length) {
      return showMessage('tracking id (guid) was not found in mail body');
    }

    var trackingId = guids[0];
    console.log('providing proof for trackingId:', trackingId);

    return getProof(trackingId, function (err, response) {
      if (err) {
        console.error('error getting proof:', err.message);
        return showMessage(err.message, event);
      }

      var proofs = response.result.proofs;
      console.log('got proofs:', proofs);

      var attachments = [];
      for (var i in proofs) {
        var proof = proofs[i];
        if (proof && proof.encryptedProof && proof.encryptedProof.sasUrl && proof.encryptedProof.documentName) {
          attachments.push({
            type: Office.MailboxEnums.AttachmentType.File,
            url: proof.encryptedProof.sasUrl,
            name: proof.encryptedProof.documentName
          })
        }
      }

      console.log('attachments: ', attachments);

      var replyText = "Please find below the requested proofs for your own validation.\r\f\r\f\r\f" + beginProofString + JSON.stringify(proofs, null, 2) + endProofString;

      var opts = {
        'htmlBody': replyText,
        'attachments': attachments
      };

      console.log('creating a reply mail with ', JSON.stringify(opts, true, 2));
      Office.context.mailbox.item.displayReplyForm(opts);

      showMessage("Proof have been added for: " + trackingId, event);
    });
  });
}

function extractGuidsFromText(text) {
  var guidRegex = new RegExp("[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}");
  var guidLength = 36;

  var index, guids = [];
  while ((index = text.search(guidRegex)) > -1) {
    var guid = text.substr(index, guidLength);
    text = text.substr(index + guidLength);
    guids.push(guid);
  }
  return guids;
}


function showMessage(message, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('ibera-notifications-id', {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: 'icon-16',
    message: message,
    persistent: false
  }, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      showMessage('Error showing a notification', event);
    }
    if (event) {
      event.completed();
    }
  });
}


/*
  MIT License:

  Permission is hereby granted, free of charge, to any person obtaining
  a copy of this software and associated documentation files (the
  'Software'), to deal in the Software without restriction, including
  without limitation the rights to use, copy, modify, merge, publish,
  distribute, sublicense, and/or sell copies of the Software, and to
  permit persons to whom the Software is furnished to do so, subject to
  the following conditions:

  The above copyright notice and this permission notice shall be
  included in all copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
  LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
  OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
  WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/