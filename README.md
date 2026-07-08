# PrimeWest External Recipient Warning

This repository hosts the static files for the PrimeWest Outlook add-in that warns users when a message is addressed to external recipients.

The add-in is served from GitHub Pages:

https://maxsilvernagel.github.io/ExternalEmailWarning/

## Key Files

- `manifest.xml` - Production Outlook add-in manifest.
- `launchevent.html` - Event runtime page loaded by Outlook on the web, new Outlook, and Outlook on Mac.
- `launchevent.js` - Event handlers for recipient changes, sensitivity label changes, and send events.
- `internal-domains.js` - Configuration for internal mail domains and the secure-send header.
- `taskpane.html` - Simple informational taskpane page.
- `assets/` - Add-in icons used by the manifest and taskpane.

## Deployment

Build the add-in from the source project, then publish the generated `dist` files to this GitHub Pages repository.

```powershell
npm run build
```

After publishing, verify these URLs return `200`:

- `https://maxsilvernagel.github.io/ExternalEmailWarning/manifest.xml`
- `https://maxsilvernagel.github.io/ExternalEmailWarning/launchevent.html`
- `https://maxsilvernagel.github.io/ExternalEmailWarning/launchevent.js`
- `https://maxsilvernagel.github.io/ExternalEmailWarning/internal-domains.js`

Use the production manifest URL when installing or updating the Outlook add-in.
