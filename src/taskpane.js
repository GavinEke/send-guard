let currentEvaluation = null;

Office.onReady(() => {
  const acknowledgement = document.getElementById("acknowledgement");
  const sendAnyway = document.getElementById("sendAnyway");
  const refresh = document.getElementById("refresh");
  const acknowledgementText = document.getElementById("acknowledgementText");

  if (acknowledgementText) {
    acknowledgementText.textContent = window.SendGuard.getConfig().acknowledgementText;
  }

  if (acknowledgement && sendAnyway) {
    acknowledgement.addEventListener("change", () => {
      sendAnyway.disabled = !acknowledgement.checked || !currentEvaluation || currentEvaluation.warnings.length === 0;
    });
  }

  if (refresh) {
    refresh.addEventListener("click", refreshWarnings);
  }

  if (sendAnyway) {
    sendAnyway.addEventListener("click", acknowledgeAndSend);
  }

  refreshWarnings();
});

async function refreshWarnings() {
  const status = document.getElementById("status");
  const warnings = document.getElementById("warnings");
  const acknowledgement = document.getElementById("acknowledgement");
  const sendAnyway = document.getElementById("sendAnyway");

  status.textContent = "Checking this email...";
  status.className = "status";
  warnings.innerHTML = "";
  acknowledgement.checked = false;
  sendAnyway.disabled = true;

  try {
    await window.SendGuard.loadConfig();
    const snapshot = await window.SendGuard.getComposeSnapshot();
    currentEvaluation = window.SendGuard.evaluateSnapshot(snapshot);

    if (currentEvaluation.warnings.length === 0) {
      status.textContent = "No data governance warnings were found for this email.";
      status.className = "status clear";
      return;
    }

    status.textContent = "This email matched one or more send warning rules.";
    status.className = "status warning";
    warnings.replaceChildren(...currentEvaluation.warnings.map(renderWarning));
  } catch (error) {
    currentEvaluation = null;
    status.textContent = "The email could not be checked. Close this pane and try sending again.";
    status.className = "status warning";
  }
}

function renderWarning(warning) {
  const container = document.createElement("article");
  const title = document.createElement("h2");
  const message = document.createElement("p");

  container.className = "warning";
  title.textContent = warning.title;
  message.textContent = warning.message;
  container.append(title, message);
  return container;
}

function acknowledgeAndSend() {
  if (!currentEvaluation || currentEvaluation.warnings.length === 0) {
    return;
  }

  window.SendGuard.setAcknowledgement(currentEvaluation);

  const item = Office.context.mailbox.item;
  if (item && typeof item.sendAsync === "function") {
    item.sendAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        showSendAgainMessage();
      }
    });
    return;
  }

  showSendAgainMessage();
}

function showSendAgainMessage() {
  const status = document.getElementById("status");
  status.textContent = "Acknowledgement saved. Select Send again to send this email.";
  status.className = "status clear";
}
