# Auto BCC — Outlook add-in

Automatically adds a BCC recipient to every message you send in Outlook
(Microsoft 365 / Exchange mailboxes).

## How it works

- Outlook fires the **`OnMessageSend`** event ("Smart Alerts") when you click Send.
- The handler in [`launchevent.js`](launchevent.js) reads your configured BCC
  address and adds it to the message, then lets the send proceed.
- The BCC address is stored in your Outlook profile via `roamingSettings` — set
  it in the **Auto BCC settings** pane on the compose ribbon. It is **not** stored
  in this repository.
- `SendMode="SoftBlock"` means the add-in never blocks a send: if it can't run,
  your message goes out normally.

## Files

| File | Purpose |
| --- | --- |
| `manifest.xml` | Add-in manifest. Sideload this into Outlook. |
| `commands.html` | Hosts the send-time event handler runtime. |
| `launchevent.js` | The `OnMessageSend` handler that adds the BCC. |
| `taskpane.html` / `taskpane.js` | Settings pane to set the BCC address. |
| `assets/` | Icons referenced by the manifest. |

## Hosting

The HTML/JS/icon files are served over HTTPS from **GitHub Pages**
(`https://danames.com/outlook-bcc/`). Only `manifest.xml` is installed
into Outlook; it points at those hosted files.

## Requirements

- Outlook on Mac (new UI) 16.62+ / current Outlook on the web / new Outlook on Windows.
- A Microsoft 365, Exchange Online, or Outlook.com mailbox (add-ins do not load
  for Gmail/IMAP accounts).

## Sideloading

1. Go to <https://aka.ms/olksideload> (signed in to your Microsoft 365 account).
2. **My add-ins → Add a custom add-in → Add from File**, then choose `manifest.xml`.
3. In a new message, open **Auto BCC settings** from the ribbon, enter your BCC
   address, and Save.
4. Send a test message and confirm the BCC arrives.
