Office.initialize = function () { };

try {
    Office.actions.associate("main", main);
} catch (e) {
    console.log("Unable to call 'Office.actions.associate'", e);
};

const PROMPT_URL = "https://andrewlamansky.github.io/Dialog-API-Bug/Prompt.html";
let clickEvent;
let dialog;

// entry point
function main(event) {
    clickEvent = event;
    openDialog();
    return;
}

function openDialog() {
    Office.context.ui.displayDialogAsync(PROMPT_URL, { height: 50, width: 50, displayInIframe: true }, dialogCallback);
}

function dialogCallback(asyncResult) {
    dialog = asyncResult.value;
    // capture message "cancel" or "move" sent from script in Prompt.html
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

    // capture event (closing Prompt.html dialog with x button)
    dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
}

function eventHandler(arg) {
    console.log(`Received event from prompt dialog.`);
    console.table(arg);
    clickEvent.completed();
}

async function messageHandler(arg) {
    console.log(`Received message "${arg.message}" from prompt dialog.`);
    dialog.close();
    switch (arg.message) {
        case "cancel":
            dialog.close();
            clickEvent.completed();
            break;
        case "move":
            let itemId = Office.context.mailbox.item.itemId;
            let folderId = "JunkEmail";
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, 
                async (result) => {
                    const token = result.value;
                    let res = await moveEmail(itemId, folderId, token);
                    console.log(`API response from request to move item to ${folderId}:`);
                    console.table(res);
                    clickEvent.completed();
                });
            break;
        default:
            break;
    }
}

async function moveEmail(itemId, folderId, token) {
    console.log(`Moving email with item ID ${itemId} to ${folderId}.`);
    const itemUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/";
    const restItemId =  Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    const junkUrl = itemUrl + restItemId + "/move";

    let res = await fetch(junkUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json;charset=utf-8',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
            "DestinationId": folderId
        })
      });
    return res;
}
