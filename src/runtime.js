Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

async function onMessageSendHandler(event) {
  try {
    await window.SendGuard.loadConfig();
    const snapshot = await window.SendGuard.getComposeSnapshot();
    const evaluation = window.SendGuard.evaluateSnapshot(snapshot);

    if (evaluation.warnings.length === 0) {
      window.SendGuard.clearAcknowledgement();
      event.completed({ allowEvent: true });
      return;
    }

    if (window.SendGuard.hasCurrentAcknowledgement(evaluation)) {
      window.SendGuard.clearAcknowledgement();
      event.completed({ allowEvent: true });
      return;
    }

    event.completed({
      allowEvent: false,
      errorMessage: window.SendGuard.smartAlertMessage(evaluation),
      commandId: "OpenPane.Button",
      contextData: JSON.stringify({ signature: evaluation.signature })
    });
  } catch (error) {
    event.completed({
      allowEvent: false,
      errorMessage:
        "The data governance send check could not complete. Try again, or contact your administrator if the problem continues."
    });
  }
}
