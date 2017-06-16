// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
// See full license at the bottom of this file.

// The initialize function is required for all add-ins.
Office.initialize = function () {
};

// TODO:  the following config values need to be set on asettings page
var config = {
  user_id : "FarmerID100"
};

const containerName = "attachments";
const beginProofString = "-----BEGIN PROOF-----";
const endProofString = "-----END PROOF-----";

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

function postProof(body, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('POST', '/api/proof');
  xhr.setRequestHeader('Content-Type', 'application/json');
  handleRequest(xhr, body, callback);
}

function getKey(keyId, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('GET', '/api/key?key_id='+ encodeURIComponent(keyId));
  xhr.setRequestHeader('Content-Type', 'application/json');
  handleRequest(xhr, null, callback);
}

function getProof(trackingId, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('GET', '/api/proof?tracking_id=' + encodeURIComponent(trackingId));
  handleRequest(xhr, null, callback);
}
function getHash(url, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('GET', '/api/hash?url=' + encodeURIComponent(url));
  handleRequest(xhr, null, callback);
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
  serviceRequest.folderName = config.user_id;
  serviceRequest.upload = upload;

  Office.context.mailbox.getCallbackTokenAsync( function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status == "succeeded") {
        serviceRequest.attachmentToken = asyncResult.value;
        var attachment;
        xhr = new XMLHttpRequest();
        // xhr.open("POST", "https://localhost:44320/api/Attachment", true);
        xhr.open("POST", "https://ibera-document-service.azurewebsites.net/api/Attachment", true);
        xhr.setRequestHeader("Content-Type", "application/json");

        // Translate the attachment details into a form easily understood by WCF.
        for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
            attachment = Office.context.mailbox.item.attachments[i];
            attachment = attachment._data$p$0 || attachment.$0_0;

            if (attachment !== undefined) {
                // I copied this line from the msdn example - not sure why first stringify and then pars the attachment
                serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
            }
        }
        handleRequest(xhr, serviceRequest, function(response) {
          callback(response, event);
        });
    }
    else {
        return showMessage("Could not get callback token: " + asyncResult.error.message, event);
    }
  });
}

function storeAttachmentsCallback(response, event) {
  if (response.error){
    return showMessage("Error: " + response.error, event);           
  }
  var trackingIds = [];
  if (response.attachmentProcessingDetails)
  {
    for (a = 0; a < response.attachmentProcessingDetails.length; a++ ){
      var ad = response.attachmentProcessingDetails[a];
      var trackingId = "id_" + ad.hash; 
      trackingIds.push(encodeURIComponent(trackingId));
      var proof = {
          tracking_id : trackingId,
          proof_to_encrypt : {
            url : ad.url,
            sas_token : ad.sasToken,
            document_name : ad.name
          },
          public_proof : {
              document_hash : ad.hash,
              creator_id : config.user_id
          }
      }; 
      postProof(proof, function(response){
        if (response.error || !response.result) {
          return showMessage(response.error, event);
        }
      });
    }
  }
  Office.context.mailbox.item.displayReplyForm(JSON.stringify(trackingIds));
  return showMessage("Attachments processed: " + JSON.stringify(trackingIds), event);
}

function getFirstAttachmentHash(event, callback){
  processAttachments(true, event, function(response, event) {
    var hash = "";
    if (response.error){
      showMessage("Error: " + response.error, event);           
    }
    if (response.attachmentProcessingDetails) {
      if (response.attachmentProcessingDetails.length>0){
        hash = response.attachmentProcessingDetails[0].hash;
      }
    }
    callback(hash);
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
                getProof(jsonProof[0].tracking_id, function(fromChain){
                  if (fromChain.error || !fromChain.result) {
                    return showMessage("error retrieving the proof from blockchain for validation - tracking_id: " + jsonProof[0].trackingId, event); 
                  }   
                  var proofFromChain = fromChain.result[0];
                  var proofToEncryptStr = JSON.stringify(jsonProof[0].encrypted_proof);
                  var hash = sha256(proofToEncryptStr);
                  if (proofFromChain.public_proof.encrypted_proof_hash == hash.toUpperCase()){
                    if (proofFromChain.public_proof.public_proof && proofFromChain.public_proof.public_proof.document_hash){
                        getFirstAttachmentHash(event, function(hash){
                          if (proofFromChain.public_proof.public_proof.document_hash == hash) {
                            return showMessage("Valid proof with attachment for tracking_id: " + jsonProof[0].tracking_id, event);
                          } 
                          else{
                            return showMessage("Valid proof BUT attachment NOT valid for tracking_id: " + jsonProof[0].tracking_id, event);
                          }
                        }); 
                    }
                    else {
                       return showMessage("Valid proof with NO attachment for tracking_id: " + jsonProof[0].tracking_id, event);                       
                    }
                  }
                  else {
                    return showMessage("NOT valid proof for tracking_id: " + jsonProof[0].tracking_id, event);               
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
	// if (Office.context.mailbox.item.subject == "ibera key request")
	{
		Office.context.mailbox.item.body.getAsync('text', {}, function (result) {
			if (result.status === Office.AsyncResultStatus.Failed) {
				return showMessage(result.error, event);
			}

			var body = result.value;
      var trackingId = body;
			getProof(trackingId, function (response) {
				if (response.error || !response.result) {
					return showMessage(response.error, event);
				}
        var proofs = response.result;
        var attachments = [];
        for (var i in proofs){
          var proof = proofs[i];
          if (proof && proof.encrypted_proof && proof.encrypted_proof.sas_token && proof.encrypted_proof.document_name) {
            attachments.push({
              type : Office.MailboxEnums.AttachmentType.File,
              url : proof.encrypted_proof.sas_token, 
              name : proof.encrypted_proof.document_name
            })
          }
        }
        var replyText = "Please find below the requested proofs for your own validation.\r\f\r\f\r\f"+ beginProofString + JSON.stringify(proofs, null, 2) + endProofString;
        Office.context.mailbox.item.displayReplyForm({
          'htmlBody' : replyText,
          'attachments' : attachments
        });
        
        showMessage("Proof have been added for: " + trackingId, event);
			});
		});
	}
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
