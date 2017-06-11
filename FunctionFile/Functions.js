// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
// See full license at the bottom of this file.

// The initialize function is required for all add-ins.
Office.initialize = function () {
};

// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody(text) {
    Office.context.mailbox.item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    Office.context.mailbox.item.body.prependAsync(
                        text,
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    Office.context.mailbox.item.body.prependAsync(
                        text,
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

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

function requestKeys(event) {
	var keyId;
	getKey(keyId, function (response) {
		if (response.error) {
			showMessage(response.error, event);
		} else {
			Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
		}
	});
}

function provideKeys(event) {
	// if (Office.context.mailbox.item.subject == "ibera key request")
	{
		showMessage("start", event);
		Office.context.mailbox.item.body.getAsync('text', {}, function (result) {
			if (result.status === Office.AsyncResultStatus.Failed) {
				showMessage(result.error, event);
				return;
			}
			var body = result.value;
			showMessage(body, event);
			getProof(body, function (response) {
				if (response.error || !response.result) {
					return showMessage(response.error, event);
				}
				var keyId = response.result[0].tracking_id;
				showMessage("retrieve key with id: " + keyId, event);
				getKey(keyId, function (response) {
					if (response.error || !response.result) {
						return showMessage(response.error, event);
					}
					try{
						setItemBody(JSON.stringify(response.result));
						showMessage("key written");
					}
					catch (ex){
						showMessage(ex.message, event);
					}
				});
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
