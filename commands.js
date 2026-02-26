// ===============================
// Domain Guard - Smart Alerts (DEBUG v2)
// ===============================

Office.initialize = function () {};
Office.onReady(function () {});

// Minimal test — block every send
function onMessageSendHandler(event) {
    event.completed({
        allowEvent: false,
        errorMessage: "Domain Guard is working! This is a test block."
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
