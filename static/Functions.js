// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
// See full license at the bottom of this file.

// The initialize function is required for all add-ins.
Office.initialize = function () {
};

setTimeout(function(){
  console.log('here');
}, 3000);

const containerName = "attachments";
const beginProofString = "-----BEGIN PROOF-----";
const endProofString = "-----END PROOF-----";

function handleRequest(xhr, body, callback) {
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if  (xhr.status === 200) {
        return callback(null, JSON.parse(xhr.responseText));
      } 
      
      console.error('status:', xhr.status);
      return callback(new Error('Request status: ' + xhr.status));
    }
  };

  xhr.onerror = function(err) {
    console.error('error:', err);
    return callback(new Error('Request error: ' + err.message));
  };

  xhr.send(body && JSON.stringify(body) || null);
}

function putProof(body, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('PUT', '/api/proof');
  xhr.setRequestHeader('Content-Type', 'application/json');
  Office.context.mailbox.getUserIdentityTokenAsync(function(userToken) {
    if (userToken.error) return callback(userToken.error);
    xhr.setRequestHeader('User-Token', userToken.value);
    handleRequest(xhr, body, callback);
  });
}

function getKey(keyId, callback) {
  var xhr = new XMLHttpRequest();
  if (keyId === decodeURIComponent(keyId)) {
    keyId = encodeURIComponent(keyId);
  }
  xhr.open('GET', '/api/key/'+ keyId);
  xhr.setRequestHeader('Content-Type', 'application/json');
  Office.context.mailbox.getUserIdentityTokenAsync(function(userToken) {
    if (userToken.error) return callback(userToken.error);
    xhr.setRequestHeader('User-Token', userToken.value);
    handleRequest(xhr, null, callback);
  });
}

function getProof(trackingId, callback) {
  var xhr = new XMLHttpRequest();
  if (trackingId === decodeURIComponent(trackingId)) {
    trackingId = encodeURIComponent(trackingId);
  }
  xhr.open('GET', '/api/proof/' + trackingId);
  Office.context.mailbox.getUserIdentityTokenAsync(function(userToken) {
    if (userToken.error) return callback(userToken.error);
    xhr.setRequestHeader('User-Token', userToken.value);
    handleRequest(xhr, null, callback);
  });
}

function getHash(url, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('GET', '/api/hash?url=' + encodeURIComponent(url));
  Office.context.mailbox.getUserIdentityTokenAsync(function(userToken) {
    if (userToken.error) return callback(userToken.error);
    xhr.setRequestHeader('User-Token', userToken.value);
    handleRequest(xhr, body, callback);
  });
}

function storeAttachments(event) {
  processAttachments(true, event, storeAttachmentsCallback);
}

function processAttachments(upload, event, callback) {
  if (Office.context.mailbox.item.attachments == undefined) {
      return showMessage("Not supported: Attachments are not supported by your Exchange server.", event);
  } else if (Office.context.mailbox.item.attachments.length == 0) {
      return showMessage("No attachments: There are no attachments on this item.", event);
  }
  var serviceRequest = {};
  serviceRequest.attachmentToken = "";
  serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
  serviceRequest.attachments = [];
  serviceRequest.containerName = containerName;
  serviceRequest.upload = upload;

  Office.context.mailbox.getCallbackTokenAsync( function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status == "succeeded") {
      serviceRequest.attachmentToken = asyncResult.value;
      var attachment;
      xhr = new XMLHttpRequest();
      xhr.open("POST", clientEnv.documentServiceUrl + "/api/Attachment", true);
      xhr.setRequestHeader("Content-Type", "application/json");
      Office.context.mailbox.getUserIdentityTokenAsync(function(userToken) {
        if (userToken.error) return callback(userToken.error);
        xhr.setRequestHeader('User-Token', userToken.value);

        // Translate the attachment details into a form easily understood by WCF.
        for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
          attachment = Office.context.mailbox.item.attachments[i];
          attachment = attachment._data$p$0 || attachment.$0_0;

          if (attachment) {
            // I copied this line from the msdn example - not sure why first stringify and then pars the attachment
            serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
          }
        }
        handleRequest(xhr, serviceRequest, function(err, response) {
          if (err) {
            return callback(err);
          }

          return callback(null, { response: response, event: event });
        });
      });
    }
    else {
      return callback(new Error("Could not get callback token: " + asyncResult.error.message));
    }
  });
}

// result: {response, event}
function storeAttachmentsCallback(err, result) {
  if (err) {
    return showMessage("Error: " + err.message, event);           
  }

  var response = result.response;
  console.log('got response', response);
  var event = result.event;

  var trackingIds = [];
  if (response.attachmentProcessingDetails)
  {
    for (a = 0; a < response.attachmentProcessingDetails.length; a++ ) {

      var ad = response.attachmentProcessingDetails[a];
      var proof = {
        proofToEncrypt : {
          url : ad.url,
          sasToken : ad.sasToken,
          documentName : ad.name
        },
        publicProof : {
          documentHash : ad.hash
        }
      }; 

      putProof(proof, function(err, response) {
        if (err) {
          return showMessage(err.message, event);
        }

        trackingIds.push(response.trackingId);

        Office.context.mailbox.item.displayReplyForm(JSON.stringify(trackingIds));
        return showMessage("Attachments processed: " + JSON.stringify(trackingIds), event);
      });
    }
  }
}

function getFirstAttachmentHash(event, callback) {

  // result: { response: response, event: event }
  processAttachments(true, event, function(err, result) {
    if (err) {
      showMessage("Error: " + err.message, event);
      return callback(err);
    }

    var response = result.response;
    var event = result.event;
  
    // response from the server
    if (response.isError) {
      console.error('error getting first attachment from server:', response.error);
      return callback(new Error(response.error));
    }

    if (!response.attachmentProcessingDetails || !response.attachmentProcessingDetails.length) {
      console.error('hash is not available');
      return callback(new Error('hash not available'));
    }

    var hash = response.attachmentProcessingDetails[0].hash;
    return callback(null, { hash: hash });

  });
}

function validateProof(event) {
		Office.context.mailbox.item.body.getAsync('text', {}, function (result) {
			if (result.status === Office.AsyncResultStatus.Failed) {
				return showMessage(result.error, event);
			}
      try {
        var body = result.value;
        if ((body.search(beginProofString) != -1) && (body.search(endProofString) != -1) ){
          var proofs = body.split(beginProofString);
          for (var i in proofs) {
            if (proofs[i].search(endProofString) != -1) {
              var proof = proofs[i].split(endProofString);
              if (proof.length >= 1) {
                var jsonProof = JSON.parse(proof[0]);
                getProof(jsonProof[0].trackingId, function(err, result) {
                  console.log('get proof from chain:', err, result);
                  if (err) {
                    return showMessage("error retrieving the proof from blockchain for validation - trackingId: " + jsonProof[0].trackingId + " error: " + err.message, event); 
                  }
                  
                  if (!result) {
                    return showMessage("error retrieving the proof from blockchain for validation - trackingId: " + jsonProof[0].trackingId, event); 
                  }

                  var proofFromChain = result.result[0];
                  var proofToEncryptStr = JSON.stringify(jsonProof[0].encryptedProof);
                  var hash = sha256(proofToEncryptStr);
                  if (proofFromChain.publicProof.encryptedProofHash == hash.toUpperCase()){
                    if (proofFromChain.publicProof.publicProof && proofFromChain.publicProof.publicProof.documentHash){
                      getFirstAttachmentHash(event, function(err, result) {
                        console.log('retrieving first attachment hash:', err, result);
                        if (err) {
                          return showMessage("error retrieving first attachment hash - trackingId: " + jsonProof[0].trackingId + " error: " + err.message, event); 
                        }

                        var hash = result.hash;
                        if (proofFromChain.publicProof.publicProof.documentHash == hash) {
                          return showMessage("Valid proof with attachment for trackingId: " + jsonProof[0].trackingId, event);
                        } 
                        else{
                          return showMessage("Valid proof BUT attachment NOT valid for trackingId: " + jsonProof[0].trackingId, event);
                        }
                      }); 
                    }
                    else {
                      return showMessage("Valid proof with NO attachment for trackingId: " + jsonProof[0].trackingId, event);                       
                    }
                  }
                  else {
                    return showMessage("NOT valid proof for trackingId: " + jsonProof[0].trackingId, event);               
                  }
                });
              }
              else {
                return showMessage("Unable to validate proof(s)", event); 
              }
            }
          }
        }
        else {
          return showMessage("No proofs to validate found in email", event); 
        }
      }
      catch(ex){
        return showMessage(ex.message, event);       
      }
    });
}

function provideProof(event) {
  Office.context.mailbox.item.body.getAsync('text', {}, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      return showMessage(result.error, event);
    }

    var body = result.value;
    var trackingId = body.trim();
    console.log('providing proof for trackingId:', trackingId);

    return getProof(trackingId, function (err, response) {
      if (err) {
        console.error('error getting proof:', err.message);
        return showMessage(err.message, event);
      }
      
      var proofs = response.result;
      console.log('got proofs:', proofs);

      var attachments = [];
      for (var i in proofs) {
        var proof = proofs[i];
        if (proof && proof.encryptedProof && proof.encryptedProof.sasToken && proof.encryptedProof.documentName) {
          attachments.push({
            type : Office.MailboxEnums.AttachmentType.File,
            url : proof.encryptedProof.sasToken, 
            name : proof.encryptedProof.documentName
          })
        }
      }

      console.log('attachments: ', attachments);

      var replyText = "Please find below the requested proofs for your own validation.\r\f\r\f\r\f"+ beginProofString + JSON.stringify(proofs, null, 2) + endProofString;
      
      var opts = {
        'htmlBody' : replyText,
        'attachments' : attachments
      };

      console.log('creating a reply mail with ', JSON.stringify(opts, true, 2));
      Office.context.mailbox.item.displayReplyForm(opts);
      
      showMessage("Proof have been added for: " + trackingId, event);
    });
  });
}

function showMessage(message, event) {
	Office.context.mailbox.item.notificationMessages.replaceAsync('ibera-notifications-id', {
		type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
		icon: 'icon-16',
		message: message,
		persistent: false
	}, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      showMessage('Error when showing a notification', event);
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
