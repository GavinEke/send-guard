(function () {
  const defaults = {
    organisationDomains: [],
    sensitiveTerms: ["client", "credit card"],
    visibleRecipientLimit: 10,
    acknowledgementText: "I confirm this email complies with my organisation's data governance policy."
  };

  window.SendGuardConfig = defaults;
  window.SendGuardConfigPromise = fetch("./config.json", { cache: "no-store" })
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
})();
