(function () {
  const ACK_KEY = "send-guard-acknowledgement";
  const CONFIG_TIMEOUT_MS = 800;
  const FIELD_TIMEOUT_MS = 1200;
  const defaultConfig = {
    organisationDomains: [],
    sensitiveTerms: ["client", "credit card"],
    visibleRecipientLimit: 10
  };
  let config = defaultConfig;

  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

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
      await loadConfig();
      const snapshot = await getComposeSnapshot();
      const evaluation = evaluateSnapshot(snapshot);

      if (evaluation.warnings.length === 0) {
        clearAcknowledgement();
        complete({ allowEvent: true });
        return;
      }

      if (hasCurrentAcknowledgement(evaluation)) {
        clearAcknowledgement();
        complete({ allowEvent: true });
        return;
      }

      complete({
        allowEvent: false,
        errorMessage: smartAlertMessage(evaluation),
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

  async function loadConfig() {
    try {
      const runtimeUrl = Office.context.urls && Office.context.urls.javascriptRuntimeUrl;
      const configUrl = runtimeUrl ? new URL("./config.json?v=1.0.2", runtimeUrl).href : "./config.json?v=1.0.2";
      const response = await withTimeout(fetch(configUrl, { cache: "no-store" }), CONFIG_TIMEOUT_MS, null);
      if (response.ok) {
        const configJson = await withTimeout(response.json(), CONFIG_TIMEOUT_MS, null);
        if (configJson) {
          config = Object.assign({}, defaultConfig, configJson);
        }
      }
    } catch (error) {
      config = defaultConfig;
    }
  }

  function getAsyncValue(source, coercionType, fallback) {
    return new Promise((resolve) => {
      let settled = false;
      const timeout = setTimeout(() => {
        settled = true;
        resolve(fallback);
      }, FIELD_TIMEOUT_MS);

      const callback = (result) => {
        if (settled) {
          return;
        }

        settled = true;
        clearTimeout(timeout);
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || fallback);
        } else {
          resolve(fallback);
        }
      };

      if (coercionType) {
        source.getAsync(coercionType, callback);
      } else {
        source.getAsync(callback);
      }
    });
  }

  async function getComposeSnapshot() {
    const item = Office.context.mailbox.item;
    const [to, cc, subject, htmlBody] = await Promise.all([
      getAsyncValue(item.to, null, []),
      getAsyncValue(item.cc, null, []),
      getAsyncValue(item.subject, null, ""),
      getAsyncValue(item.body, Office.CoercionType.Html, "")
    ]);

    return {
      to: Array.isArray(to) ? to : [],
      cc: Array.isArray(cc) ? cc : [],
      subject: subject || "",
      bodyText: htmlToText(htmlBody || "")
    };
  }

  function htmlToText(html) {
    if (typeof document !== "undefined") {
      const element = document.createElement("div");
      element.innerHTML = html;
      return (element.textContent || element.innerText || "").replace(/\s+/g, " ").trim();
    }

    return String(html)
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function evaluateSnapshot(snapshot) {
    const internalDomains = getOrganisationDomains();
    const visibleEmails = [...snapshot.to, ...snapshot.cc].map(recipientEmail).filter(Boolean);
    const externalEmails = visibleEmails.filter((email) => isExternalEmail(email, internalDomains));
    const sensitiveTerms = findSensitiveTerms(`${snapshot.subject}\n${snapshot.bodyText}`);
    const warnings = [];

    if (visibleEmails.length > config.visibleRecipientLimit) {
      warnings.push({
        code: "too-many-visible-recipients",
        title: "Large visible recipient list",
        message: `This email has ${visibleEmails.length} recipients in To/CC. Consider moving recipients to BCC before sending.`
      });
    }

    if (sensitiveTerms.length > 0 && externalEmails.length > 0) {
      warnings.push({
        code: "sensitive-external",
        title: "Sensitive content sent externally",
        message: `Sensitive term(s) found: ${sensitiveTerms.join(", ")}. External recipient(s): ${externalEmails.join(", ")}.`
      });
    }

    return {
      warnings,
      signature: createSignature({
        recipients: visibleEmails.sort(),
        subject: snapshot.subject,
        bodyText: snapshot.bodyText,
        warningCodes: warnings.map((warning) => warning.code).sort()
      })
    };
  }

  function getOrganisationDomains() {
    const configured = config.organisationDomains.map(normaliseDomain).filter(Boolean);
    if (configured.length > 0) {
      return configured;
    }

    const senderEmail = Office.context.mailbox.userProfile.emailAddress || "";
    const senderDomain = normaliseDomain(senderEmail.split("@").pop() || "");
    return senderDomain ? [senderDomain] : [];
  }

  function recipientEmail(recipient) {
    return normaliseEmail(recipient && (recipient.emailAddress || recipient.displayName || ""));
  }

  function normaliseEmail(email) {
    return String(email || "").trim().toLowerCase();
  }

  function normaliseDomain(domain) {
    return String(domain || "").trim().toLowerCase();
  }

  function isExternalEmail(email, internalDomains) {
    const domain = normaliseDomain(String(email).split("@").pop() || "");
    return Boolean(domain) && !internalDomains.some((internal) => domain === internal || domain.endsWith(`.${internal}`));
  }

  function findSensitiveTerms(text) {
    const normalisedText = String(text || "").toLowerCase();
    return config.sensitiveTerms.filter((term) => normalisedText.includes(String(term).toLowerCase()));
  }

  function createSignature(value) {
    const serialized = JSON.stringify(value);
    let hash = 0;
    for (let index = 0; index < serialized.length; index += 1) {
      hash = (hash << 5) - hash + serialized.charCodeAt(index);
      hash |= 0;
    }
    return `${serialized.length}:${Math.abs(hash)}`;
  }

  function getAcknowledgement() {
    try {
      return JSON.parse(localStorage.getItem(ACK_KEY) || "null");
    } catch (error) {
      return null;
    }
  }

  function clearAcknowledgement() {
    localStorage.removeItem(ACK_KEY);
  }

  function hasCurrentAcknowledgement(evaluation) {
    const acknowledgement = getAcknowledgement();
    return Boolean(acknowledgement && acknowledgement.signature === evaluation.signature);
  }

  function smartAlertMessage(evaluation) {
    const lines = evaluation.warnings.map((warning) => `- ${warning.title}: ${warning.message}`);
    return [
      "Review this email before sending.",
      ...lines,
      "Select Review send warnings to open Send Guard and acknowledge your organisation's data governance policy."
    ]
      .join("\n")
      .slice(0, 500);
  }

  function withTimeout(promise, timeoutMs, fallback) {
    return new Promise((resolve) => {
      const timeout = setTimeout(() => resolve(fallback), timeoutMs);
      promise
        .then((value) => {
          clearTimeout(timeout);
          resolve(value);
        })
        .catch(() => {
          clearTimeout(timeout);
          resolve(fallback);
        });
    });
  }
})();
