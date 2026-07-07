/*
 * Auto BCC settings pane.
 *
 * Writes the BCC address to two local stores the send-time handler can read:
 *  1. roamingSettings — per-mailbox storage in the user's Outlook/Exchange
 *     profile (primary).
 *  2. localStorage — same-origin fallback for runtimes where roamingSettings
 *     isn't readable at send time.
 * Neither store ever leaves the user's Outlook profile / machine.
 */

var SETTING_KEY = "autoBccAddress";
var EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) {
    return;
  }

  var current = readCurrent();
  if (current) {
    document.getElementById("bcc").value = current;
    setStatus("Currently BCC'ing " + current, "ok");
  }

  document.getElementById("save").onclick = save;
  document.getElementById("clear").onclick = clearAddress;
});

function readCurrent() {
  var value = "";
  try {
    value = (Office.context.roamingSettings.get(SETTING_KEY) || "").trim();
  } catch (e) {}
  if (!value) {
    try {
      value = (localStorage.getItem(SETTING_KEY) || "").trim();
    } catch (e) {}
  }
  return value;
}

function setStatus(message, kind) {
  var el = document.getElementById("status");
  el.textContent = message;
  el.className = kind || "";
}

function writeLocalStorage(value) {
  try {
    if (value) {
      localStorage.setItem(SETTING_KEY, value);
    } else {
      localStorage.removeItem(SETTING_KEY);
    }
    return true;
  } catch (e) {
    return false;
  }
}

function persist(value, successMessage) {
  // Fallback store first — it's synchronous and can't fail silently later.
  var localOk = writeLocalStorage(value);

  // Primary store: roamingSettings.
  try {
    if (value) {
      Office.context.roamingSettings.set(SETTING_KEY, value);
    } else {
      Office.context.roamingSettings.remove(SETTING_KEY);
    }
    Office.context.roamingSettings.saveAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setStatus(successMessage, "ok");
      } else if (localOk) {
        setStatus(successMessage + " (saved locally; profile sync failed)", "ok");
      } else {
        setStatus("Couldn't save: " + result.error.message, "err");
      }
    });
  } catch (e) {
    if (localOk) {
      setStatus(successMessage + " (saved locally)", "ok");
    } else {
      setStatus("Couldn't save the address.", "err");
    }
  }
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
