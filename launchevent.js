/*
 * Auto BCC — adds a configured BCC recipient to every message on send.
 *
 * The BCC address is NOT stored here. It is read at send time from
 * roamingSettings (the user's Outlook/Exchange profile), with a
 * localStorage fallback — both set via the settings task pane. That
 * keeps the address out of this public file entirely.
 *
 * Design rule: this handler must NEVER block a send. Every path — success,
 * failure, thrown exception, or stalled callback — ends in
 * event.completed({ allowEvent: true }).
 */

var SETTING_KEY = "autoBccAddress";
var WATCHDOG_MS = 15000; // absolute ceiling on how long we hold a send

function getConfiguredAddress() {
  var address = "";

  // Primary store: roamingSettings (per-mailbox, synced by Exchange).
  try {
    var rs = Office && Office.context && Office.context.roamingSettings;
    if (rs) {
      address = (rs.get(SETTING_KEY) || "").trim();
    }
  } catch (e) {
    // roamingSettings can be unavailable in some event runtimes — fall through.
  }

  // Fallback store: localStorage (same origin as the settings pane).
  if (!address) {
    try {
      if (typeof localStorage !== "undefined" && localStorage) {
        address = (localStorage.getItem(SETTING_KEY) || "").trim();
      }
    } catch (e) {
      // localStorage may be blocked; nothing more we can do.
    }
  }

  return address;
}

function onMessageSendHandler(event) {
  var done = false;
  var watchdog = null;

  function finish() {
    if (done) {
      return;
    }
    done = true;
    if (watchdog !== null) {
      try { clearTimeout(watchdog); } catch (e) {}
    }
    try {
      event.completed({ allowEvent: true });
    } catch (e) {
      // Nothing left to do — never rethrow out of the handler.
    }
  }

  // Watchdog: if any callback below never fires, release the send anyway.
  try {
    if (typeof setTimeout === "function") {
      watchdog = setTimeout(finish, WATCHDOG_MS);
    }
  } catch (e) {}

  try {
    var address = getConfiguredAddress();

    // No address configured — send the message untouched.
    if (!address) {
      finish();
      return;
    }

    var item = Office.context.mailbox.item;
    if (!item || !item.bcc) {
      finish();
      return;
    }

    // Don't add a duplicate if this address is already on the Bcc line.
    item.bcc.getAsync(function (getResult) {
      try {
        var alreadyPresent = false;

        if (
          getResult.status === Office.AsyncResultStatus.Succeeded &&
          getResult.value
        ) {
          for (var i = 0; i < getResult.value.length; i++) {
            var existing = (getResult.value[i].emailAddress || "").toLowerCase();
            if (existing === address.toLowerCase()) {
              alreadyPresent = true;
              break;
            }
          }
        }

        if (alreadyPresent) {
          finish();
          return;
        }

        item.bcc.addAsync([address], function (addResult) {
          if (addResult.status === Office.AsyncResultStatus.Failed) {
            try {
              console.error("Auto BCC: failed to add BCC — " + addResult.error.message);
            } catch (e) {}
          }
          finish();
        });
      } catch (e) {
        finish();
      }
    });
  } catch (e) {
    finish();
  }
}

// Required: map the manifest's handler name to this function.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
