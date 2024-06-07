// Office onReady function to initialize the add-in
Office.onReady(function (info) {
    // Ensure the DOM is ready before initializing the add-in
    $(document).ready(function () {
        // Add event handler for ItemSend event
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemSend,
            function (args) {
                // Get the current item being sent
                var item = args.item;

                // Specify the text you want to search for in the subject line
                var searchTerms = ["[HDP-stl_legal.FID", "[HDP-troy_legal.FID", "[HDP-dc_legal.FID", "[HDP-firm_admin.FID"];

                // Check if any of the search terms exist in the subject
                var containsSearchTerm = searchTerms.some(function (term) {
                    return item.subject.indexOf(term) !== -1;
                });

                // Prompt the user if any search term is found
                if (containsSearchTerm) {
                    // Display a confirmation dialog
                    if (!confirm("***WARNING***\n\nThis message appears to have a luggage tag and might be filed to the DM automatically if you send this email. Are you sure you wish to send this message?")) {
                        // Cancel sending the item
                        args.completed({ allowEvent: false });
                    }
                }

                // Continue sending the item
                args.completed({ allowEvent: true });
            }
        );
    });
});
