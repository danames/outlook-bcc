/*
 * Auto BCC — adds a configured BCC recipient to every message on send.
 *
 * The BCC address is NOT stored here. It is read at send time from
 * roamingSettings, which lives in the user's Outlook/Exchange profile
 * (set via the settings task pane). That keeps the address out of this
 * public file entirely.
 */

var SETTING_KEY = "autoBccAddress";

function onMessageSendHandler(event) {
  var address = (Office.context.roamingSettings.get(SETTING_KEY) || "").trim();

  // No address configured yet — send the message untouched.
  if (!address) {
    event.completed({ allowEvent: true });
    return;
  }

  var item = Office.context.mailbox.item;

  // Don't add a duplicate if this address is already on the Bcc line.
  item.bcc.getAsync({ asyncContext: event }, function (getResult) {
    var ev = getResult.asyncContext;
    var alreadyPresent = false;

    if (getResult.status === Office.AsyncResultStatus.Succeeded && getResult.value) {
      for (var i = 0; i < getResult.value.length; i++) {
        var existing = (getResult.value[i].emailAddress || "").toLowerCase();
        if (existing === address.toLowerCase()) {
          alreadyPresent = true;
          break;
        }
      }
    }

    if (alreadyPresent) {
      ev.completed({ allowEvent: true });
      return;
    }

    item.bcc.addAsync([address], { asyncContext: ev }, function (addResult) {
      if (addResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Auto BCC: failed to add BCC — " + addResult.error.message);
      }
      // "Send anyway" behavior: let the message go regardless of whether
      // adding the BCC succeeded, so the add-in never blocks a send.
      addResult.asyncContext.completed({ allowEvent: true });
    });
  });
}

// Required: map the manifest's handler name to this function.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
