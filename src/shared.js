(function () {
  const ACK_KEY = "send-guard-acknowledgement";

  function getConfig() {
    return window.SendGuardConfig || {};
  }

  function loadConfig() {
    return window.SendGuardConfigPromise || Promise.resolve(getConfig());
  }

  function getMailboxItem() {
    return Office.context.mailbox.item;
  }

  function getSenderDomain() {
    const email = Office.context.mailbox.userProfile.emailAddress || "";
    return normaliseDomain(email.split("@").pop() || "");
  }

  function getOrganisationDomains() {
    const configured = getConfig().organisationDomains || [];
    const domains = configured.map(normaliseDomain).filter(Boolean);
    if (domains.length > 0) {
      return domains;
    }

    const senderDomain = getSenderDomain();
    return senderDomain ? [senderDomain] : [];
  }

  function normaliseDomain(domain) {
    return String(domain || "").trim().toLowerCase();
  }

  function normaliseEmail(email) {
    return String(email || "").trim().toLowerCase();
  }

  function recipientEmail(recipient) {
    return normaliseEmail(recipient && (recipient.emailAddress || recipient.displayName || ""));
  }

  function getAsyncValue(source, coercionType) {
    return new Promise((resolve, reject) => {
      const callback = (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          reject(result.error);
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
    const item = getMailboxItem();
    const [to, cc, subject, htmlBody] = await Promise.all([
      getAsyncValue(item.to),
      getAsyncValue(item.cc),
      getAsyncValue(item.subject),
      getAsyncValue(item.body, Office.CoercionType.Html)
    ]);

    return {
      to: Array.isArray(to) ? to : [],
      cc: Array.isArray(cc) ? cc : [],
      subject: subject || "",
      bodyText: htmlToText(htmlBody || "")
    };
  }

  function htmlToText(html) {
    const element = document.createElement("div");
    element.innerHTML = html;
    return (element.textContent || element.innerText || "").replace(/\s+/g, " ").trim();
  }

  function isExternalEmail(email, internalDomains) {
    const domain = normaliseDomain(String(email).split("@").pop() || "");
    if (!domain) {
      return false;
    }

    return !internalDomains.some((internalDomain) => {
      return domain === internalDomain || domain.endsWith(`.${internalDomain}`);
    });
  }

  function findSensitiveTerms(text) {
    const configuredTerms = getConfig().sensitiveTerms || [];
    const normalisedText = String(text || "").toLowerCase();
    return configuredTerms.filter((term) => {
      const normalisedTerm = String(term || "").trim().toLowerCase();
      return normalisedTerm && normalisedText.includes(normalisedTerm);
    });
  }

  function evaluateSnapshot(snapshot) {
    const internalDomains = getOrganisationDomains();
    const visibleRecipients = [...snapshot.to, ...snapshot.cc];
    const visibleEmails = visibleRecipients.map(recipientEmail).filter(Boolean);
    const externalEmails = visibleEmails.filter((email) => isExternalEmail(email, internalDomains));
    const combinedText = `${snapshot.subject}\n${snapshot.bodyText}`;
    const sensitiveTerms = findSensitiveTerms(combinedText);
    const visibleRecipientLimit = Number(getConfig().visibleRecipientLimit || 10);

    const warnings = [];
    if (visibleEmails.length > visibleRecipientLimit) {
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
      visibleRecipientCount: visibleEmails.length,
      externalEmails,
      sensitiveTerms,
      signature: createSignature({
        recipients: visibleEmails.sort(),
        subject: snapshot.subject,
        bodyText: snapshot.bodyText,
        warningCodes: warnings.map((warning) => warning.code).sort()
      })
    };
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

  function setAcknowledgement(evaluation) {
    localStorage.setItem(
      ACK_KEY,
      JSON.stringify({
        signature: evaluation.signature,
        acknowledgedAt: new Date().toISOString()
      })
    );
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
      "Use Review send warnings to acknowledge your organisation's data governance policy."
    ]
      .join("\n")
      .slice(0, 500);
  }

  window.SendGuard = {
    clearAcknowledgement,
    evaluateSnapshot,
    getAcknowledgement,
    getComposeSnapshot,
    getConfig,
    hasCurrentAcknowledgement,
    loadConfig,
    setAcknowledgement,
    smartAlertMessage
  };
})();
