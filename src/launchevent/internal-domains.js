/* global window */

window.primeWestExternalWarningConfig = {
  internalDomains: ["primewest.org"],

  // Optional diagnostics only. Sensitivity labels do not skip the secure-send prompt.
  secureSensitivityLabelIds: [],
  secureSensitivityLabelNames: [],

  // Configure a Purview / Exchange Online mail flow rule to encrypt when this header is present.
  secureSendHeaderName: "X-PrimeWest-Secure-Send",
  secureSendHeaderValue: "true",
};
