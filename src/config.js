(function () {
  const CONFIG_TIMEOUT_MS = 800;
  const defaults = {
    organisationDomains: [],
    sensitiveTerms: ["client", "credit card"],
    visibleRecipientLimit: 10,
    acknowledgementText: "I confirm this email complies with my organisation's data governance policy."
  };

  window.SendGuardConfig = defaults;
  window.SendGuardConfigPromise = withTimeout(fetch("./config.json", { cache: "no-store" }), CONFIG_TIMEOUT_MS, null)
    .then((response) => {
      if (!response.ok) {
        return defaults;
      }
      return response.json();
    })
    .then((config) => {
      window.SendGuardConfig = { ...defaults, ...config };
      return window.SendGuardConfig;
    })
    .catch(() => defaults);

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
