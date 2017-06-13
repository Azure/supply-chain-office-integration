// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
// See full license at the bottom of this file.

// The initialize function is required for all add-ins.
Office.initialize = function () {
};

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

function getKey(keyId, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('GET', '/api/key?key_id='+keyId);
  xhr.setRequestHeader('Content-Type', 'application/json');
  handleRequest(xhr, null, callback);
}

function getProof(trackingId, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('GET', '/api/proof?tracking_id=' + trackingId);
  handleRequest(xhr, null, callback);
}



function storeAttachments(event){

  if (Office.context.mailbox.item.attachments == undefined) {
    var testButton = document.getElementById("testButton");
      testButton.onclick = "";
      showMessage("Not supported: Attachments are not supported by your Exchange server.", event);
  } else if (Office.context.mailbox.item.attachments.length == 0) {
      var testButton = document.getElementById("testButton");
      testButton.onclick = "";
      showMessage("No attachments: There are no attachments on this item.", event);
  } else {
    var serviceRequest = new Object();
    serviceRequest.attachmentToken = "";
    serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
    serviceRequest.attachments = new Array();
    serviceRequest.containerName = "attachments";
    serviceRequest.folderName = "username";
  }
  showMessage("Start", event);

  Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
}

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status == "succeeded") {
        serviceRequest.attachmentToken = asyncResult.value;
        makeServiceRequest();
    }
    else {
        showMessage("Could not get callback token: " + asyncResult.error.message);
    }
}

function makeServiceRequest() {
    showMessage("makeServiceRequest");
    var attachment;
    xhr = new XMLHttpRequest();

    // Update the URL to point to your service location.
    // xhr.open("POST", "https://localhost:44320/api/Attachment", true);
    xhr.open("POST", "https://ibera-document-service/api/Attachment", true);
   

    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");

    // Translate the attachment details into a form easily understood by WCF.
    for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
        attachment = Office.context.mailbox.item.attachments[i];
        attachment = attachment._data$p$0 || attachment.$0_0;

        if (attachment !== undefined) {
            serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
        }
    }
    handleRequest(xhr, serviceRequest, showResponse);
};

// Shows the service response.
function showResponse(response) {
    showMessage("Attachments processed: " + response.attachmentsProcessed);
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
                  if (proofFromChain.public_proof.encrypted_proof_hash == sha256(jsonProof[0].encrypted_proof)){
                    return showMessage("Valid proof for tracking_id: " + jsonProof[0].tracking_id, event);  
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
        var replyText = "Please find below the requested proofs for your own validation.\r\f\r\f\r\f"+ beginProofString + JSON.stringify(response.result, null, 2) + endProofString;
        Office.context.mailbox.item.displayReplyForm(replyText);
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
