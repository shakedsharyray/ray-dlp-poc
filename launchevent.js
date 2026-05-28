var BLOCK_TERM = "bamba";

function onMessageSendHandler(event) {
  var done = false;
  function finish(opts) {
    if (done) return;
    done = true;
    event.completed(opts);
  }

  var timeout = setTimeout(function () {
    finish({
      allowEvent: false,
      errorMessage:
        "Ray DLP could not classify the message in time. Send blocked.",
    });
  }, 4000);

  Office.context.mailbox.item.subject.getAsync(function (subjResult) {
    var subject = ((subjResult && subjResult.value) || "").toLowerCase();
    if (subject.indexOf(BLOCK_TERM) !== -1) {
      clearTimeout(timeout);
      finish({
        allowEvent: false,
        errorMessage:
          "Blocked by Ray DLP: the subject contains '" + BLOCK_TERM + "'.",
      });
      return;
    }

    Office.context.mailbox.item.body.getAsync("text", function (bodyResult) {
      clearTimeout(timeout);
      if (!bodyResult || bodyResult.status === "failed") {
        finish({
          allowEvent: false,
          errorMessage: "Ray DLP could not read the message body. Send blocked.",
        });
        return;
      }
      var body = (bodyResult.value || "").toLowerCase();
      if (body.indexOf(BLOCK_TERM) !== -1) {
        finish({
          allowEvent: false,
          errorMessage:
            "Blocked by Ray DLP: this message contains '" +
            BLOCK_TERM +
            "'. Please remove it and try again.",
        });
      } else {
        finish({ allowEvent: true });
      }
    });
  });
}

Office.onReady(function () {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
