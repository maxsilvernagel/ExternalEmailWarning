/*
 * Copyright (c) Chris Folkert. All rights reserved.
 */

/**
 * In classic Outlook on Windows, when the event handler runs, code in Office.onReady() or Office.initialize isn't run.
 * Add any startup logic needed by handlers to the event handler itself.
 */
Office.onReady();

/**
 * Configure the SMTP domains that should be treated as internal.
 * Add every corporate mail domain that should not trigger the warning.
 * @type {string[]}
 */
const INTERNAL_DOMAINS = normalizeInternalDomains(
  window.primeWestExternalWarningConfig && window.primeWestExternalWarningConfig.internalDomains
);

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
  "This message is addressed to at least one recipient outside of your organization.\n\n" +
  "Review the recipient list and confirm whether this message should be sent through your secure mail process before continuing.";

/**
 * Smart Alerts message shown when recipient validation cannot complete in time.
 * @type {string}
 */
const SEND_TIMEOUT_MESSAGE =
  "Recipient validation could not be completed in time.\n\n" +
  "Select Don't send, wait a moment, and try again. If the problem continues, verify your recipients and Outlook add-in connectivity before sending.";

function normalizeInternalDomains(domains) {
  return (domains || [])
    .map((domain) => (domain || "").trim().toLowerCase())
    .filter((domain, index, allDomains) => domain && allDomains.indexOf(domain) === index);
}

function onMessageRecipientsChangedHandler(event) {
  getAllRecipients((error, recipients) => {
    if (error) {
      console.log(error);
      event.completed();
      return;
    }

    const externalRecipients = getExternalRecipients(recipients);
    logRecipientDiagnostics(recipients, externalRecipients);

    if (externalRecipients.length > 0) {
      showExternalRecipientNotification(event);
      return;
    }

    clearExternalRecipientNotification(event);
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

  getAllRecipients((error, recipients) => {
    clearTimeout(timeoutId);

    if (error) {
      console.log(error);
      complete({ allowEvent: true });
      return;
    }

    const externalRecipients = getExternalRecipients(recipients);
    logRecipientDiagnostics(recipients, externalRecipients);

    if (externalRecipients.length > 0) {
      complete({ allowEvent: false, errorMessage: SEND_WARNING_MESSAGE });
      return;
    }

    complete({ allowEvent: true });
  });
}

function createSafeSendCompletion(event) {
  let completed = false;

  return (options) => {
    if (completed) {
      return;
    }

    completed = true;
    event.completed(options);
  };
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
      callback(`Unable to get the recipients from the ${fieldName} field. Error: ${result.error.message}`, []);
      return;
    }

    callback(null, result.value || []);
  });
}

function getExternalRecipients(recipients) {
  return recipients.filter((recipient) => isExternalRecipient(recipient.emailAddress));
}

function isExternalRecipient(emailAddress) {
  const domain = getDomainFromAddress(emailAddress);
  if (!domain) {
    return false;
  }

  return !getEffectiveInternalDomains().some((internalDomain) => isSameOrChildDomain(domain, internalDomain));
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

function logRecipientDiagnostics(recipients, externalRecipients) {
  const recipientSummary = recipients.map((recipient) => ({
    displayName: recipient.displayName,
    emailAddress: recipient.emailAddress,
    domain: getDomainFromAddress(recipient.emailAddress),
  }));

  const externalSummary = externalRecipients.map((recipient) => recipient.emailAddress);

  console.log("Configured internal domains:", INTERNAL_DOMAINS);
  console.log("Effective internal domains:", getEffectiveInternalDomains());
  console.log("Recipient summary:", recipientSummary);
  console.log("External recipients:", externalSummary);
}

function showExternalRecipientNotification(event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    EXTERNAL_WARNING_NOTIFICATION_ID,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "External recipient detected. Review whether this message should be sent through your secure mail process before sending.",
      icon: "Icon.80x80",
      persistent: true,
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log(`Unable to show the external recipient warning. Error: ${result.error.message}`);
      }

      event.completed();
    }
  );
}

function clearExternalRecipientNotification(event) {
  Office.context.mailbox.item.notificationMessages.removeAsync(EXTERNAL_WARNING_NOTIFICATION_ID, () => {
    event.completed();
  });
}

Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);