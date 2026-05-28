function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Text,
    function (result) {
      if (result.status === Office.AsyncResult.ErrorStatus) {
        event.completed({
          allowEvent: false,
          errorMessage: "Ray DLP could not inspect this message. Send blocked.",
        });
        return;
      }
      const body = (result.value || "").toLowerCase();
      if (body.indexOf("bamba") !== -1) {
        event.completed({
          allowEvent: false,
          errorMessage:
            "Blocked by Ray DLP: this message contains the prohibited term 'bamba'. " +
            "Please remove it and try again.",
        });
      } else {
        event.completed({ allowEvent: true });
      }
    }
  );
}

Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
