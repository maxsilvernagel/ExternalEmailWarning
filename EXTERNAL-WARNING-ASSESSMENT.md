# Outlook New Add-in Assessment

## Current state

This repository is a Microsoft sample for an **event-based Outlook add-in**. It is a valid starting point for New Outlook because it already uses the add-in model that replaces VSTO for this scenario.

The sample currently does **not** implement an "external email warning" flow. It implements a different policy:

- watch compose events
- inspect recipients and attachments
- set a sensitivity label
- show a Smart Alerts prompt on send

That means the important part is already here: the sample proves the event-driven add-in model, manifest structure, and JavaScript runtime pattern needed for New Outlook.

## What the sample is using

- `manifest-localhost.xml`: local development manifest pointing to `https://localhost:3000`
- `manifest.xml`: hosted sample manifest pointing to GitHub
- `src/launchevent/launchevent.js`: event handlers for compose/send events
- `src/launchevent/launchevent.html`: runtime page used by New Outlook and OWA
- `package.json`: Node/Webpack tooling used to build and sideload the add-in locally

## What is blocked right now on this machine

The current environment does not have Node.js or npm available on `PATH`, and `node_modules` is not present.

That means the sample cannot currently be started locally with:

```powershell
npm install
npm start
```

until Node.js is installed and dependencies are restored.

## How this maps to the external warning requirement

The closest fit in this sample is the `OnMessageSend` event with `SendMode="PromptUser"` and the recipient-change events.

For an external recipient warning add-in, the likely shape is:

1. On recipient change, inspect `To`, `Cc`, and `Bcc`.
2. Determine whether any recipient is external.
3. On send, if at least one external recipient exists, show a Smart Alerts prompt.
4. If the user chooses to stop, cancel send.
5. If the user chooses to continue, allow send.

This is much closer to a policy/prompt add-in than to the sensitivity label behavior in this sample.

## Important design question

The phrase "send secure" is the main missing requirement.

There are two very different implementations:

1. Warn only
   The add-in shows a prompt when external recipients are present, but does not alter delivery.

2. Enforce or initiate secure delivery
   The add-in changes the message before send or hands off to another secure-mail mechanism.

Option 2 depends on the organization's actual secure-mail product and process. Office.js can show prompts and modify some compose fields, but "send secure" is not a built-in generic Outlook action. That behavior must be defined against the real mail security solution.

## Recommended next step

Use this sample as the base, but replace the sensitivity-label logic with external-recipient detection.

Initial implementation target:

- remove attachment and sensitivity label logic
- keep event-based activation
- keep `OnMessageRecipientsChanged`
- keep `OnMessageSend` with `PromptUser`
- add config for internal domains such as the company SMTP domains
- prompt if any recipient domain is not internal

## Suggested Jira direction

This request is best tracked as a small spike / proof-of-concept under the existing Outlook New add-in work rather than as a general support task.

Suggested scope for the spike:

- confirm New Outlook add-in architecture for this use case
- get Microsoft sample running locally
- modify sample into external-recipient warning proof of concept
- document limitations around "send secure" integration

## Practical setup checklist

- install Node.js LTS
- run `npm install`
- run `npm start`
- sideload `manifest-localhost.xml` into New Outlook
- validate that launch events fire in compose mode
- then replace sample logic in `src/launchevent/launchevent.js`

## Likely code areas to change first

- `src/launchevent/launchevent.js`
- `manifest-localhost.xml`
- `manifest.xml`
- `src/taskpane/taskpane.html`
- `src/taskpane/taskpane.js`

## Bottom line

The Microsoft sample is a good technical starting point for New Outlook.

The immediate blocker is environment setup, not the add-in model itself. After Node.js is available, this sample should be able to serve as the base for an external-recipient warning proof of concept.
