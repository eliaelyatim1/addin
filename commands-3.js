function autoInsertSignature(event) {
    event.completed();
}

Office.actions.associate("autoInsertSignature", autoInsertSignature);
