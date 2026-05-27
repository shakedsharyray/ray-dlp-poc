function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage:
      "Blocked by Ray DLP (test mode). " +
      "All outbound mail is currently dropped. " +
      "Contact your administrator if you need to send.",
  });
}

Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
