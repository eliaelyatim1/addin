// commands.js - Desktop Outlook event-based runtime (JavaScript only)
// Reads the pre-built signature HTML from roamingSettings.
// The taskpane stores the signature after successful Graph API auth.

// Required: ensures the Office runtime is initialized before associate() is called
Office.onReady(function() {});

function autoInsertSignature(event) {
    try {
        var sig = Office.context.roamingSettings.get("sig_html_he");
        if (!sig) {
            Office.context.mailbox.item.body.setSelectedDataAsync(
                "[חתימה לא נמצאה - פתח את ה-taskpane ולחץ הכנס חתימה פעם אחת]",
                { coercionType: Office.CoercionType.Text },
                function() { event.completed(); }
            );
            return;
        }

        Office.context.mailbox.item.body.setSignatureAsync(
            sig,
            { coercionType: Office.CoercionType.Html },
            function(result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    Office.context.mailbox.item.body.setSelectedDataAsync(
                        "[שגיאה בהכנסת חתימה: " + result.error.message + "]",
                        { coercionType: Office.CoercionType.Text },
                        function() { event.completed(); }
                    );
                } else {
                    event.completed();
                }
            }
        );
    } catch (err) {
        Office.context.mailbox.item.body.setSelectedDataAsync(
            "[שגיאה כללית: " + err.message + "]",
            { coercionType: Office.CoercionType.Text },
            function() { event.completed(); }
        );
    }
}

// Must be at global scope (not inside Office.onReady) for Classic Outlook Desktop
Office.actions.associate("autoInsertSignature", autoInsertSignature);
