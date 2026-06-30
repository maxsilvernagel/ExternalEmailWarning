/*
 * Copyright (c) Chris Folkert. All rights reserved.
 */

/**
 * In classic Outlook on Windows, when the event handler runs, code in Office.onReady() or Office.initialize isn't run.
 * Add any startup logic needed by handlers to the event handler itself.
 */
/* global Office, window, console, setTimeout, clearTimeout */

Office.onReady();

const ADD_IN_CONFIG =
  (typeof window !== "undefined" && window.primeWestExternalWarningConfig) || {};

/**
 * Configure the SMTP domains that should be treated as internal.
 * Add every corporate mail domain that should not trigger the warning.
 * @type {string[]}
 */
const INTERNAL_DOMAINS = normalizeInternalDomains(ADD_IN_CONFIG.internalDomains);

/**
 * Optional sensitivity label IDs used for diagnostics only.
 * IDs are preferred in production because label display names can change.
 * @type {string[]}
 */
const SECURE_SENSITIVITY_LABEL_IDS = normalizeStringList(
  ADD_IN_CONFIG.secureSensitivityLabelIds || ADD_IN_CONFIG.secureSensitivityLabelId
);

/**
 * Optional sensitivity label names used for diagnostics only.
 * These labels do not skip the external-recipient send prompt.
 * @type {string[]}
 */
const SECURE_SENSITIVITY_LABEL_NAMES = normalizeStringList(
  ADD_IN_CONFIG.secureSensitivityLabelNames || ADD_IN_CONFIG.secureSensitivityLabelName || []
);

/**
 * Custom internet header that asks Exchange Online / Purview mail flow rules to secure the message.
 * This avoids changing subject/body content that external recipients and filters can see.
 * @type {string}
 */
const SECURE_SEND_HEADER_NAME = normalizeHeaderName(
  ADD_IN_CONFIG.secureSendHeaderName || "X-PrimeWest-Secure-Send"
);

/**
 * Header value matched by the secure mail flow rule.
 * @type {string}
 */
const SECURE_SEND_HEADER_VALUE = normalizeConfiguredHeaderValue(
  ADD_IN_CONFIG.secureSendHeaderValue || "true"
);

/**
 * Manifest command ID used by the Smart Alerts action button.
 * @type {string}
 */
const SECURE_SEND_COMMAND_ID = "sendSecurelyButton";
const SECURE_SEND_BUTTON_LABEL = "Send Securely";
const SECURE_SEND_NOTIFICATION_ID = "secure-send-marker-added";

let cachedSecureSensitivityLabelIds = null;

/**
 * Notification ID used for the compose infobar message.
 * @type {string}
 */
const EXTERNAL_WARNING_NOTIFICATION_ID = "external-recipient-warning";

/**
 * The maximum amount of time to wait before failing with a controlled Smart Alerts message.
 * @type {number}
 */
const SEND_CHECK_TIMEOUT_MS = 4000;

/**
 * Smart Alerts message shown when the user tries to send a message to an external recipient.
 * @type {string}
 */
const SEND_WARNING_MESSAGE =
  "External recipient detected.\n\n" +
  "Should this message be sent securely?\n\n" +
  "Select Send Securely to add the Purview secure-mail marker, or Send Anyway to send without that marker.";

/**
 * Notification shown after the user chooses the secure-send action from Smart Alerts or the ribbon.
 * @type {string}
 */

const SECURE_SEND_MANUAL_SEND_MESSAGE =
  "Secure-send marker added. Select Send again to route this message through the Purview secure mail rule.";

const SECURE_SEND_FAILED_TO_SEND_MESSAGE =
  "Secure-send marker added, but Outlook could not automatically send the message. Select Send again to continue.";

const SECURE_SEND_MARK_FAILED_MESSAGE =
  "Unable to add the secure-send marker. Try again, or contact IT before sending sensitive information externally.";

/**
 * Smart Alerts message shown when recipient validation cannot complete in time.
 * @type {string}
 */
const SEND_TIMEOUT_MESSAGE =
  "Recipient validation could not be completed in time.\n\n" +
  "Select Don't send, wait a moment, and try again. If the problem continues, verify your recipients and Outlook add-in connectivity before sending.";

function normalizeStringList(values) {
  const list = Array.isArray(values) ? values : [values];

  return list
    .map((value) => (value || "").trim().toLowerCase())
    .filter((value, index, allValues) => value && allValues.indexOf(value) === index);
}

function normalizeInternalDomains(domains) {
  return normalizeStringList(domains);
}

function normalizeHeaderName(value) {
  const headerName = (value || "").trim();
  return headerName || "X-PrimeWest-Secure-Send";
}

function normalizeConfiguredHeaderValue(value) {
  const headerValue = (value || "").trim();
  return (headerValue || "true").toLowerCase();
}

function normalizeObservedHeaderValue(value) {
  return (value || "").trim().toLowerCase();
}

function onMessageRecipientsChangedHandler(event) {
  refreshExternalRecipientNotification(event);
}

function onSensitivityLabelChangedHandler(event) {
  refreshExternalRecipientNotification(event);
}

function refreshExternalRecipientNotification(event) {
  evaluateMessageSecurity((error, messageSecurity) => {
    if (error) {
      console.log(error);
      event.completed();
      return;
    }

    logRecipientDiagnostics(messageSecurity);

    if (messageSecurity.externalRecipients.length > 0) {
      showExternalRecipientNotification(event, messageSecurity);
      return;
    }

    removeSecureSendMarker(() => clearExternalRecipientNotification(event));
  });
}

function onMessageSendHandler(event) {
  const complete = createSafeSendCompletion(event);

  const timeoutId = setTimeout(() => {
    complete({
      allowEvent: false,
      errorMessage: SEND_TIMEOUT_MESSAGE,
    });
  }, SEND_CHECK_TIMEOUT_MS);

  evaluateMessageSecurity((error, messageSecurity) => {
    clearTimeout(timeoutId);

    if (error) {
      console.log(error);
      complete({ allowEvent: true });
      return;
    }

    logRecipientDiagnostics(messageSecurity);

    if (messageSecurity.shouldWarn) {
      complete(createSecureSendPromptOptions(messageSecurity));
      return;
    }

    complete({ allowEvent: true });
  });
}

function sendSecurelyHandler(event) {
  const complete = createSafeSendCompletion(event);

  markMessageForSecureSend((error) => {
    if (error) {
      console.log(error);
      showSecureSendNotification(SECURE_SEND_MARK_FAILED_MESSAGE, () => complete());
      return;
    }

    Office.context.mailbox.item.notificationMessages.removeAsync(
      EXTERNAL_WARNING_NOTIFICATION_ID,
      () => {
        sendCurrentMessage((sendError) => {
          if (sendError) {
            console.log(sendError);
            showSecureSendNotification(SECURE_SEND_FAILED_TO_SEND_MESSAGE, () => complete());
            return;
          }

          complete();
        });
      }
    );
  });
}

function createSecureSendPromptOptions(messageSecurity) {
  return {
    allowEvent: false,
    errorMessage: SEND_WARNING_MESSAGE,
    cancelLabel: SECURE_SEND_BUTTON_LABEL,
    commandId: SECURE_SEND_COMMAND_ID,
    contextData: JSON.stringify({
      action: "mark-for-secure-send",
      externalRecipientCount: messageSecurity.externalRecipients.length,
      secureSendHeaderName: SECURE_SEND_HEADER_NAME,
    }),
  };
}

function createSafeSendCompletion(event) {
  let completed = false;

  return (options) => {
    if (completed) {
      return;
    }

    completed = true;
    if (event && typeof event.completed === "function") {
      event.completed(options);
    }
  };
}

function evaluateMessageSecurity(callback) {
  getAllRecipients((error, recipients) => {
    if (error) {
      callback(error);
      return;
    }

    const externalRecipients = getExternalRecipients(recipients);
    if (externalRecipients.length === 0) {
      callback(null, {
        recipients,
        externalRecipients,
        isMarkedSecure: false,
        isMarkedForSecureSend: false,
        isMarkedWithSecureSensitivityLabel: false,
        shouldWarn: false,
        sensitivityLabelStatus: "no-external-recipients",
      });
      return;
    }

    getMessageSecurityStatus((securityStatusError, messageSecurityStatus) => {
      if (securityStatusError) {
        console.log(securityStatusError);
      }

      callback(null, {
        recipients,
        externalRecipients,
        isMarkedSecure: messageSecurityStatus.isMarkedForSecureSend,
        isMarkedForSecureSend: messageSecurityStatus.isMarkedForSecureSend,
        isMarkedWithSecureSensitivityLabel:
          messageSecurityStatus.isMarkedWithSecureSensitivityLabel,
        shouldWarn: !messageSecurityStatus.isMarkedForSecureSend,
        sensitivityLabelStatus: messageSecurityStatus.sensitivityLabelStatus,
      });
    });
  });
}

function getAllRecipients(callback) {
  getRecipientsForField(Office.context.mailbox.item.to, "To", (toError, toRecipients) => {
    if (toError) {
      callback(toError, []);
      return;
    }

    getRecipientsForField(Office.context.mailbox.item.cc, "Cc", (ccError, ccRecipients) => {
      if (ccError) {
        callback(ccError, []);
        return;
      }

      getRecipientsForField(Office.context.mailbox.item.bcc, "Bcc", (bccError, bccRecipients) => {
        if (bccError) {
          callback(bccError, []);
          return;
        }

        callback(null, [...toRecipients, ...ccRecipients, ...bccRecipients]);
      });
    });
  });
}

function getRecipientsForField(recipientField, fieldName, callback) {
  recipientField.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      callback(
        `Unable to get the recipients from the ${fieldName} field. Error: ${result.error.message}`,
        []
      );
      return;
    }

    callback(null, result.value || []);
  });
}

function getMessageSecurityStatus(callback) {
  isMessageMarkedForSecureSend((secureSendHeaderError, isMarkedForSecureSend) => {
    if (secureSendHeaderError) {
      console.log(secureSendHeaderError);
    }

    if (isMarkedForSecureSend) {
      callback(null, {
        isMarkedForSecureSend: true,
        isMarkedWithSecureSensitivityLabel: false,
        sensitivityLabelStatus: "matched-secure-send-header",
      });
      return;
    }

    isMessageMarkedWithSecureSensitivityLabel(
      (sensitivityLabelError, isMarkedWithSecureSensitivityLabel, sensitivityLabelStatus) => {
        if (sensitivityLabelError) {
          console.log(sensitivityLabelError);
        }

        callback(null, {
          isMarkedForSecureSend: false,
          isMarkedWithSecureSensitivityLabel,
          sensitivityLabelStatus,
        });
      }
    );
  });
}

function isMessageMarkedForSecureSend(callback) {
  if (!canUseInternetHeaders()) {
    callback("Internet header APIs are not available in this Outlook context.", false);
    return;
  }

  Office.context.mailbox.item.internetHeaders.getAsync([SECURE_SEND_HEADER_NAME], (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      callback(`Unable to get the secure-send header. Error: ${result.error.message}`, false);
      return;
    }

    const currentHeaderValue = getInternetHeaderValue(result.value, SECURE_SEND_HEADER_NAME);
    callback(null, normalizeObservedHeaderValue(currentHeaderValue) === SECURE_SEND_HEADER_VALUE);
  });
}

function markMessageForSecureSend(callback) {
  if (!canUseInternetHeaders()) {
    callback("Internet header APIs are not available in this Outlook context.");
    return;
  }

  const headers = {};
  headers[SECURE_SEND_HEADER_NAME] = SECURE_SEND_HEADER_VALUE;

  Office.context.mailbox.item.internetHeaders.setAsync(headers, (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      callback(`Unable to add the secure-send header. Error: ${result.error.message}`);
      return;
    }

    callback(null);
  });
}

function sendCurrentMessage(callback) {
  if (!canSendCurrentMessage()) {
    showSecureSendNotification(SECURE_SEND_MANUAL_SEND_MESSAGE, () => callback(null));
    return;
  }

  Office.context.mailbox.item.sendAsync((result) => {
    if (result && result.status === Office.AsyncResultStatus.Failed) {
      callback(`Unable to automatically send the message. Error: ${result.error.message}`);
      return;
    }

    callback(null);
  });
}

function canSendCurrentMessage() {
  return (
    Office.context &&
    Office.context.mailbox &&
    Office.context.mailbox.item &&
    typeof Office.context.mailbox.item.sendAsync === "function"
  );
}

function removeSecureSendMarker(callback) {
  if (!canUseInternetHeaders()) {
    callback();
    return;
  }

  Office.context.mailbox.item.internetHeaders.removeAsync([SECURE_SEND_HEADER_NAME], () => {
    callback();
  });
}

function canUseInternetHeaders() {
  return (
    Office.context &&
    Office.context.mailbox &&
    Office.context.mailbox.item &&
    Office.context.mailbox.item.internetHeaders &&
    typeof Office.context.mailbox.item.internetHeaders.setAsync === "function" &&
    typeof Office.context.mailbox.item.internetHeaders.getAsync === "function"
  );
}

function getInternetHeaderValue(headers, headerName) {
  if (!headers) {
    return "";
  }

  const exactHeaderValue = headers[headerName];
  if (exactHeaderValue) {
    return exactHeaderValue;
  }

  const normalizedHeaderName = headerName.toLowerCase();
  const matchingHeaderName = Object.keys(headers).filter(
    (name) => name.toLowerCase() === normalizedHeaderName
  )[0];
  return matchingHeaderName ? headers[matchingHeaderName] : "";
}

function isMessageMarkedWithSecureSensitivityLabel(callback) {
  if (
    !Office.context ||
    !Office.context.sensitivityLabelsCatalog ||
    !Office.context.mailbox ||
    !Office.context.mailbox.item ||
    !Office.context.mailbox.item.sensitivityLabel
  ) {
    callback(
      "Sensitivity label APIs are not available in this Outlook context.",
      false,
      "api-unavailable"
    );
    return;
  }

  Office.context.sensitivityLabelsCatalog.getIsEnabledAsync((catalogResult) => {
    if (catalogResult.status === Office.AsyncResultStatus.Failed) {
      callback(
        `Unable to verify whether the sensitivity label catalog is enabled. Error: ${catalogResult.error.message}`,
        false,
        "catalog-status-error"
      );
      return;
    }

    if (catalogResult.value !== true) {
      callback(null, false, "catalog-disabled");
      return;
    }

    Office.context.mailbox.item.sensitivityLabel.getAsync((labelResult) => {
      if (labelResult.status === Office.AsyncResultStatus.Failed) {
        callback(
          `Unable to get the message sensitivity label. Error: ${labelResult.error.message}`,
          false,
          "current-label-error"
        );
        return;
      }

      const currentLabelId = normalizeSensitivityLabelId(labelResult.value);
      if (!currentLabelId) {
        callback(null, false, "no-current-label");
        return;
      }

      if (SECURE_SENSITIVITY_LABEL_IDS.indexOf(currentLabelId) !== -1) {
        callback(null, true, "matched-configured-label-id");
        return;
      }

      resolveSecureSensitivityLabelIds((catalogError, secureSensitivityLabelIds) => {
        if (catalogError) {
          callback(catalogError, false, "secure-label-resolution-error");
          return;
        }

        const isSecure = secureSensitivityLabelIds.indexOf(currentLabelId) !== -1;
        callback(null, isSecure, isSecure ? "matched-catalog-label" : "label-not-secure");
      });
    });
  });
}

function resolveSecureSensitivityLabelIds(callback) {
  if (cachedSecureSensitivityLabelIds) {
    callback(null, cachedSecureSensitivityLabelIds);
    return;
  }

  Office.context.sensitivityLabelsCatalog.getAsync((catalogResult) => {
    if (catalogResult.status === Office.AsyncResultStatus.Failed) {
      callback(
        `Unable to retrieve the sensitivity label catalog. Error: ${catalogResult.error.message}`,
        []
      );
      return;
    }

    const secureLabelIds = [...SECURE_SENSITIVITY_LABEL_IDS];
    collectSecureSensitivityLabelIds(catalogResult.value || [], secureLabelIds);
    cachedSecureSensitivityLabelIds = normalizeStringList(secureLabelIds);

    callback(null, cachedSecureSensitivityLabelIds);
  });
}

function collectSecureSensitivityLabelIds(labels, secureLabelIds) {
  labels.forEach((label) => {
    if (isConfiguredSecureSensitivityLabel(label)) {
      collectSensitivityLabelAndChildrenIds(label, secureLabelIds);
      return;
    }

    collectSecureSensitivityLabelIds(label.children || [], secureLabelIds);
  });
}

function collectSensitivityLabelAndChildrenIds(label, secureLabelIds) {
  const labelId = normalizeSensitivityLabelId(label.id);
  if (labelId) {
    secureLabelIds.push(labelId);
  }

  (label.children || []).forEach((childLabel) =>
    collectSensitivityLabelAndChildrenIds(childLabel, secureLabelIds)
  );
}

function isConfiguredSecureSensitivityLabel(label) {
  const labelId = normalizeSensitivityLabelId(label.id);
  const labelName = (label.name || "").trim().toLowerCase();

  return (
    (labelId && SECURE_SENSITIVITY_LABEL_IDS.indexOf(labelId) !== -1) ||
    (labelName && SECURE_SENSITIVITY_LABEL_NAMES.indexOf(labelName) !== -1)
  );
}

function normalizeSensitivityLabelId(labelId) {
  return (labelId || "").trim().toLowerCase();
}

function getExternalRecipients(recipients) {
  return recipients.filter((recipient) => isExternalRecipient(recipient.emailAddress));
}

function isExternalRecipient(emailAddress) {
  const domain = getDomainFromAddress(emailAddress);
  if (!domain) {
    return false;
  }

  return !getEffectiveInternalDomains().some((internalDomain) =>
    isSameOrChildDomain(domain, internalDomain)
  );
}

function getDomainFromAddress(emailAddress) {
  if (typeof emailAddress !== "string") {
    return "";
  }

  const normalizedEmailAddress = emailAddress.trim().toLowerCase();
  const atSymbolIndex = normalizedEmailAddress.lastIndexOf("@");
  if (atSymbolIndex === -1) {
    return "";
  }

  return normalizedEmailAddress.slice(atSymbolIndex + 1);
}

function getMailboxDomain() {
  const mailboxEmailAddress =
    Office.context &&
    Office.context.mailbox &&
    Office.context.mailbox.userProfile &&
    Office.context.mailbox.userProfile.emailAddress;

  return getDomainFromAddress(mailboxEmailAddress);
}

function getEffectiveInternalDomains() {
  return normalizeInternalDomains([...INTERNAL_DOMAINS, getMailboxDomain()]);
}

function isSameOrChildDomain(domain, internalDomain) {
  const normalizedInternalDomain = (internalDomain || "").trim().toLowerCase();
  if (!normalizedInternalDomain) {
    return false;
  }

  return domain === normalizedInternalDomain || domain.endsWith(`.${normalizedInternalDomain}`);
}

function logRecipientDiagnostics(messageSecurity) {
  const recipientSummary = messageSecurity.recipients.map((recipient) => ({
    displayName: recipient.displayName,
    emailAddress: recipient.emailAddress,
    domain: getDomainFromAddress(recipient.emailAddress),
  }));

  const externalSummary = messageSecurity.externalRecipients.map(
    (recipient) => recipient.emailAddress
  );

  console.log("Configured internal domains:", INTERNAL_DOMAINS);
  console.log("Effective internal domains:", getEffectiveInternalDomains());
  console.log("Configured secure sensitivity label names:", SECURE_SENSITIVITY_LABEL_NAMES);
  console.log("Configured secure sensitivity label IDs:", SECURE_SENSITIVITY_LABEL_IDS);
  console.log("Secure send header:", SECURE_SEND_HEADER_NAME);
  console.log("Sensitivity label status:", messageSecurity.sensitivityLabelStatus);
  console.log(
    "Message marked with secure sensitivity label:",
    messageSecurity.isMarkedWithSecureSensitivityLabel
  );
  console.log("Message marked for secure send:", messageSecurity.isMarkedForSecureSend);
  console.log("Recipient summary:", recipientSummary);
  console.log("External recipients:", externalSummary);
}

function showExternalRecipientNotification(event, messageSecurity) {
  const message = messageSecurity.isMarkedForSecureSend
    ? "External recipient detected. This message is already marked for secure delivery."
    : "External recipient detected. When you send, choose whether to mark this message for Purview secure delivery.";

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    EXTERNAL_WARNING_NOTIFICATION_ID,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.80x80",
      persistent: true,
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(
          `Unable to show the external recipient warning. Error: ${result.error.message}`
        );
      }

      event.completed();
    }
  );
}

function showSecureSendNotification(message, callback) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    SECURE_SEND_NOTIFICATION_ID,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.80x80",
      persistent: true,
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(`Unable to show the secure-send notification. Error: ${result.error.message}`);
      }

      callback();
    }
  );
}

function clearExternalRecipientNotification(event) {
  Office.context.mailbox.item.notificationMessages.removeAsync(
    EXTERNAL_WARNING_NOTIFICATION_ID,
    () => {
      event.completed();
    }
  );
}

Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
Office.actions.associate("onSensitivityLabelChangedHandler", onSensitivityLabelChangedHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("sendSecurelyHandler", sendSecurelyHandler);
