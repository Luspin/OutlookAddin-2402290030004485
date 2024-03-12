Office.onReady((info) => {
    console.log("Office.onReady called");
    if (info.host === Office.HostType.Outlook) {
        // document.getElementById("helloButton").onclick = sayHello;
        // document.getElementById("displayDialogAsyncButton").onclick = openDialog;
        // document.getElementById("openBrowserWindowButton").onclick = openBrowserWindow;
        // document.getElementById("syncMessageButton").onclick = syncMessage;
        // document.getElementById("sendMessageButtonGraph").onclick = sendMessageGraph;

        console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.13")));

    }
});

/**
 * Writes 'Hello world!' to a new message Subject and Body. # UPDATE
 */
function sayHello() {
    console.log("Saying hello");

    Office.context.mailbox.item.body.setAsync(
        "Hello world!",
        {
            coercionType: "html", // Write text as HTML
        },

        // Callback method to check that setAsync succeeded
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }
        }
    );



    // sendGETRequest();
}

function sendGETRequest() {

    var xhr = new XMLHttpRequest();

    xhr.onload = function () {
        if (xhr.status === 200) {
            // Process the response data
            console.log(xhr.responseText);
        } else {
            // Handle errors
            console.error('Request failed. Status: ', xhr.status);
        }
    };

    xhr.open('GET', 'https://oam.lusp.in:8443/')

    xhr.send();
}

let dialog; // Declare dialog as global for use in later functions.

function openDialog() {
    console.log("Opening dialog");

    Office.context.ui.displayDialogAsync('https://luspin.github.io/OutlookAddin/myDialog.html', { height: 60, width: 30, promptBeforeOpen: false },
        function (asyncResult) {
            dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                processMessage(arg);
            });
        }
    );
}

let accessToken;

function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message.slice(1, -1).replace(/\\"/g, '"'));
    console.log(messageFromDialog);

    if (messageFromDialog.messageType === "dialogClosed") {
        console.log("Dialog closed");
        document.getElementById("dialogResultText").innerHTML = "Result: " + messageFromDialog.messageType;
        dialog.close();
    }

    if (messageFromDialog.messageType === "userAuthenticated") {
        console.log("user Authenticated");
        document.getElementById("dialogResultText").innerHTML = "Hello: " + messageFromDialog.displayName;
        console.log(messageFromDialog.accessToken);
        accessToken = messageFromDialog.accessToken;
        dialog.close();
    }
}

function openBrowserWindow() {
    Office.context.ui.openBrowserWindow("https://www.google.com");
}

let savedMailId;

async function syncMessage() {
    Office.context.mailbox.item.saveAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`saveAsync succeeded, itemId is ${result.value}`);
        } else {
            console.error(`saveAsync failed with message ${result.error.message}`);
        }

        const restId = Office.context.mailbox.convertToRestId(result.value, Office.MailboxEnums.RestVersion.v2_0);

        console.log("REST item ID: " + restId);
        document.getElementById("syncMessageIdLabel").innerHTML = "Synced message ID: " + restId;
        savedMailId = restId;
    });
}

async function sendMessageGraph() {
    // https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API/Using_Fetch
    // https://learn.microsoft.com/en-us/graph/api/message-send?view=graph-rest-1.0&tabs=http
    console.log("Sending message using Graph API");

    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/messages/' + savedMailId + '/send', {
            method: "POST",
            headers: {
                'Authorization': 'Bearer ' + accessToken
            }
        });

        const messageSendStatusDetails = response.json(); // This returns a Promise
        // Wait for the JSON promise to resolve
        const messageSendStatus = await messageSendStatusDetails;

        document.getElementById("sendMessageStatusLabel").innerHTML = "Message sent status: " + messageSendStatus;
    } catch (error) {
        // undefined
        document.getElementById("errorMessage").innerHTML = "Error: " + error.message;
    }
}