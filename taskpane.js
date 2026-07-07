/*
 * Auto BCC settings pane.
 * Reads/writes the BCC address to roamingSettings (per-mailbox storage in the
 * user's Outlook/Exchange profile). The send-time handler reads the same value.
 */

var SETTING_KEY = "autoBccAddress";
var EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  var current = Office.context.roamingSettings.get(SETTING_KEY);
  if (current) {
    document.getElementById("bcc").value = current;
    setStatus("Currently BCC'ing " + current, "ok");
  }

  document.getElementById("save").onclick = save;
  document.getElementById("clear").onclick = clearAddress;
});

function setStatus(message, kind) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = kind || "";
}

function persist(value, successMessage) {
  Office.context.roamingSettings.set(SETTING_KEY, value);
  Office.context.roamingSettings.saveAsync(function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      setStatus(successMessage, "ok");
    } else {
      setStatus("Couldn't save: " + result.error.message, "err");
    }
  });
}

function save() {
  var value = document.getElementById("bcc").value.trim();

  if (!value) {
    setStatus("Enter an address, or choose \"Turn off\".", "err");
    return;
  }
  if (!EMAIL_RE.test(value)) {
    setStatus("That doesn't look like a valid email address.", "err");
    return;
  }

  persist(value, "Saved. Every message you send will now BCC " + value + ".");
}

function clearAddress() {
  document.getElementById("bcc").value = "";
  persist("", "Turned off. No BCC will be added.");
}
