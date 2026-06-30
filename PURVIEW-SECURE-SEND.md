# Purview secure-send marker

This add-in marks an external message for secure delivery by adding a custom internet header, then automatically sending the message when the user selects **Send Securely** from the Smart Alerts dialog.

Configured header:

```text
X-PrimeWest-Secure-Send: true
```

Why a header: it avoids changing the subject or body, so the marker is not visible in normal recipient mail clients and is less likely to look like spammy content.

## Tenant setup

Create an Exchange Online mail flow rule for Microsoft Purview Message Encryption:

1. Go to Exchange admin center > Mail flow > Rules.
2. Create a rule for messages where the sender is inside the organization.
3. Add a condition for recipients outside the organization.
4. Add a condition where a message header includes `X-PrimeWest-Secure-Send` and the header value includes `true`.
5. Add the action to apply Office 365 Message Encryption and rights protection, then choose the appropriate Encrypt-Only or custom RMS template.
6. Optionally remove the `X-PrimeWest-Secure-Send` header after the encryption action, if your mail flow policy allows that action order.

Test with a small pilot group before broad deployment. The add-in only adds the marker; the mail flow rule is what actually encrypts the message. Sensitivity labels are only diagnostic in this add-in and do not automatically choose secure delivery for the user.

## How to verify the header

The header is not visible in the normal message body or subject. To confirm it was stamped, send a test message to an external mailbox you can inspect, then open the message source / message details in that receiving mailbox and search for:

```text
X-PrimeWest-Secure-Send: true
```

Good places to check:

- Outlook on the web: open the received test message, use the message details / view source option, then search for the header name.
- Outlook classic desktop: open the received test message, open message properties, then inspect Internet headers.
- Microsoft 365 admin tooling: use message trace or Microsoft Graph against the received message if your admin permissions allow reading `internetMessageHeaders`.

For a quick add-in-side check before send, the compose banner changes to say the message is already marked for secure delivery after the secure-send marker is present. If the client does not support programmatic send from the function command, the add-in falls back to adding the marker and asking the user to select **Send** again.
