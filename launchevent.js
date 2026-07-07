/* DEBUG BUILD — minimal handler for bisection. Completes immediately. */
function onMessageSendHandler(event) {
  event.completed({ allowEvent: true });
}
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
