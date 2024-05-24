Office.onReady(function () {
    document.getElementById("sendButton").onclick = checkBeforeSending;
});

function checkBeforeSending() {
    Office.context.mailbox.item.subject.getAsync(function (result) {
        var subject = result.value;
        var specificText = ".FID"; // Text to search for in the subject line
        var prompt = "***WARNING***\n\nThis message has a luggage tag and will be filed to the DM automatically if you send this email. Are you sure you wish to send this message?";
        
        if (subject.indexOf(specificText) !== -1) {
            if (confirm(prompt)) {
                Office.context.mailbox.item.send();
            }
        } else {
            Office.context.mailbox.item.send();
        }
    });
}
