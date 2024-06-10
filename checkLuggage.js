console.log("checkLuggage.js is being executed...");

Office.initialize = function (reason) {
    console.log("Office.initialize is being called...");
    
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, function (args) {
        console.log("ItemSend event handler is invoked...");
        
        var item = args.get_item();
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            var subject = item.subject;
            console.log("Message subject: " + subject);
            
            // Specify the text you want to search for in the subject line
            var searchTerms = ["[HDP-stl_legal.FID", "[HDP-troy_legal.FID", "[HDP-dc_legal.FID", "[HDP-firm_admin.FID"];

            // Check if any of the search terms exist in the subject
            var containsSearchTerm = searchTerms.some(function (term) {
                return subject.indexOf(term) !== -1;
            });

            if (containsSearchTerm) {
                console.log("Subject contains a search term. Displaying dialog...");
                
                // Show a confirmation dialog
                Office.context.ui.displayDialogAsync(
                    'https://jmharn.github.io/Check4LuggageTa/confirmation.html',
                    { height: 200, width: 400 },
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            var dialog = result.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
                                if (args.message === 'cancel') {
                                    console.log("User chose to cancel the send operation.");
                                    args.completed();
                                    item.notificationMessages.addAsync('warning', prompt, { type: 'errorMessage' });
                                    args.completed({ cancel: true });
                                } else {
                                    console.log("User chose to send the message.");
                                    args.completed();
                                }
                            });
                        } else {
                            console.error("Failed to display the dialog: " + result.error.message);
                        }
                    }
                );
            } else {
                console.log("Subject does not contain a search term.");
            }
        }
    });
};
