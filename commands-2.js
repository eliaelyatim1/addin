// commands.js - Desktop Outlook event-based runtime (JavaScript only)
// Reads the pre-built signature HTML from roamingSettings.
// The taskpane stores the signature after successful Graph API auth.

// Required: ensures the Office runtime is initialized before associate() is called
Office.onReady(function() {});

function autoInsertSignature(event) {
    try {
        var sig = Office.context.roamingSettings.get("sig_html_he");
        if (!sig) {
            event.completed();
            return;
        }

        Office.context.mailbox.item.body.setSignatureAsync(
            sig,
            { coercionType: Office.CoercionType.Html },
            function() { event.completed(); }
        );
    } catch (err) {
        event.completed();
    }
}

// Must be at global scope (not inside Office.onReady) for Classic Outlook Desktop
Office.actions.associate("autoInsertSignature", autoInsertSignature);
