console.log("MessageRead.js is being executed...");

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, onItemSend);
    }
});

function onItemSend(eventArgs) {
    console.log("onItemSend function is invoked...");
    const item = eventArgs.item;
    item.subject.getAsync({ asyncContext: eventArgs }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const subject = asyncResult.value;
            const searchTerms = ["[HDP-stl_legal.FID", "[HDP-troy_legal.FID", "[HDP-dc_legal.FID", "[HDP-firm_admin.FID"];
            const containsSearchTerm = searchTerms.some(function (term) {
                return subject.includes(term);
            });

            if (containsSearchTerm) {
                const promptMessage = "***WARNING***\n\nThis message appears to have a luggage tag and might be filed to the DM automatically if you send this email. Are you sure you wish to send this message?";
                Office.context.ui.displayDialogAsync('https://jmharn.github.io/Check4LuggageTag/Prompt.html', { height: 30, width: 20 }, function (result) {
                    let dialog = result.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                        if (arg.message === 'no') {
                            eventArgs.completed({ allowEvent: false });
                        } else {
                            eventArgs.completed({ allowEvent: true });
                        }
                        dialog.close();
                    });
                });
            } else {
                eventArgs.completed({ allowEvent: true });
            }
        } else {
            console.error("Failed to get subject: " + asyncResult.error.message);
            eventArgs.completed({ allowEvent: true });
        }
    });
}
