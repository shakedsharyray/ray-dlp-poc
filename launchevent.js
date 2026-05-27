// Outlook fires OnMessageSend when the user clicks Send.
// We always block, with a dialog message. This is the entire DLP logic.

function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage:
      "Blocked by Ray DLP (test mode). " +
      "All outbound mail is currently dropped. " +
      "Contact your administrator if you need to send.",
  });
}

// Required: register the function name so Office can invoke it.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
