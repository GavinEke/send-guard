async function onMessageSendHandler(event) {
  let completed = false;
  const watchdog = setTimeout(() => {
    complete({
      allowEvent: false,
      errorMessage:
        "Send Guard needs you to review this email before sending. Select Review send warnings to open Send Guard.",
      commandId: "OpenPane.Button",
      contextData: JSON.stringify({ reason: "timeout" })
    });
  }, 3000);

  function complete(options) {
    if (completed) {
      return;
    }

    completed = true;
    clearTimeout(watchdog);
    event.completed(options);
  }

  try {
    await window.SendGuard.loadConfig();
    const snapshot = await window.SendGuard.getComposeSnapshot();
    const evaluation = window.SendGuard.evaluateSnapshot(snapshot);

    if (evaluation.warnings.length === 0) {
      window.SendGuard.clearAcknowledgement();
      complete({ allowEvent: true });
      return;
    }

    if (window.SendGuard.hasCurrentAcknowledgement(evaluation)) {
      window.SendGuard.clearAcknowledgement();
      complete({ allowEvent: true });
      return;
    }

    complete({
      allowEvent: false,
      errorMessage: window.SendGuard.smartAlertMessage(evaluation),
      commandId: "OpenPane.Button",
      contextData: JSON.stringify({ signature: evaluation.signature })
    });
  } catch (error) {
    complete({
      allowEvent: false,
      errorMessage:
        "Send Guard needs you to review this email before sending. Select Review send warnings to open Send Guard.",
      commandId: "OpenPane.Button",
      contextData: JSON.stringify({ reason: "error" })
    });
  }
}

Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
