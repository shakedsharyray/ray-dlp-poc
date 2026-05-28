// Ray DLP — Smart Alerts handler that delegates classification to the
// FastAPI classifier service. Edit policy in the service, not here.

var CLASSIFIER_URL = "https://pecan-shudder-nutty.ngrok-free.dev/classify";
var TIMEOUT_MS = 4500;                 // stay under OWA's 5s long-running threshold
var ATTACHMENT_SIZE_LIMIT = 512 * 1024; // skip fetching content for >512KB attachments

function onMessageSendHandler(event) {
  var done = false;
  function finish(opts) {
    if (done) return;
    done = true;
    event.completed(opts);
  }

  var timer = setTimeout(function () {
    finish({
      allowEvent: false,
      errorMessage: "Ray DLP could not classify the message in time. Send blocked.",
    });
  }, TIMEOUT_MS);

  var item = Office.context.mailbox.item;
  var profile = Office.context.mailbox.userProfile || {};
  var data = {
    sender: profile.emailAddress || null,
    subject: "",
    body: "",
    to: [],
    cc: [],
    attachments: [],
    itemId: item.itemId || null,
    timestamp: new Date().toISOString(),
  };

  function postClassify() {
    fetch(CLASSIFIER_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "ngrok-skip-browser-warning": "1",
      },
      body: JSON.stringify(data),
    })
      .then(function (res) { return res.ok ? res.json() : null; })
      .then(function (verdict) {
        clearTimeout(timer);
        if (!verdict) {
          finish({
            allowEvent: false,
            errorMessage: "Ray DLP service returned an invalid response. Send blocked.",
          });
          return;
        }
        if (verdict.action === "block") {
          finish({
            allowEvent: false,
            errorMessage: verdict.reason || "Blocked by Ray DLP.",
          });
        } else {
          finish({ allowEvent: true });
        }
      })
      .catch(function () {
        clearTimeout(timer);
        finish({
          allowEvent: false,
          errorMessage: "Ray DLP service unreachable. Send blocked.",
        });
      });
  }

  function collectAttachments(callback) {
    item.getAttachmentsAsync(function (ar) {
      var atts = (ar && ar.value) || [];
      if (atts.length === 0) {
        callback([]);
        return;
      }
      var out = new Array(atts.length);
      var pending = atts.length;
      atts.forEach(function (att, idx) {
        var meta = {
          id: att.id,
          name: att.name,
          contentType: att.contentType,
          size: att.size,
          isInline: !!att.isInline,
          attachmentType: att.attachmentType,
          content: null,
          format: null,
        };
        // Skip content for inline images and oversized attachments
        if (att.isInline || (att.size && att.size > ATTACHMENT_SIZE_LIMIT)) {
          out[idx] = meta;
          if (--pending === 0) callback(out);
          return;
        }
        item.getAttachmentContentAsync(att.id, function (cr) {
          if (cr && cr.status !== "failed" && cr.value) {
            meta.content = cr.value.content;
            meta.format = cr.value.format;
          }
          out[idx] = meta;
          if (--pending === 0) callback(out);
        });
      });
    });
  }

  item.subject.getAsync(function (r) {
    if (r && r.value) data.subject = r.value;
    item.body.getAsync("text", function (r) {
      if (r && r.value) data.body = r.value;
      item.to.getAsync(function (r) {
        if (r && r.value) data.to = r.value;
        item.cc.getAsync(function (r) {
          if (r && r.value) data.cc = r.value;
          collectAttachments(function (atts) {
            data.attachments = atts;
            postClassify();
          });
        });
      });
    });
  });
}

Office.onReady(function () {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
