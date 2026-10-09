# 06 — Frontend

The client decides which server call happens next and keeps all retry state; the server is
stateless functions.

## Pages

Apps Script serves one HTML file per page, so each page in `client/pages/` is assembled from
fragments:

- `include('name')` inlines a fragment as written; `includeTemplate('name')` evaluates its
  scriptlets first. Only the error scripts need the template form, to inline the error codes.
- Order matters: styles in `<head>`, then sections, then the page's globals script, then fragment
  scripts. A function exists only once its `<script>` block has run, so code that runs on load can
  only use earlier blocks.
- Each page ends by loading Google's `api.js`, whose `onload` starts the page.

| Page | Folder | Starts with |
| --- | --- | --- |
| `update.html` | `client/update/` | `gapiLoaded()` |
| `get_started.html` | `client/get_started/` | `getStartedGapiLoaded()` |
| `save_file.html` | `client/save_file/` | `saveFileGapiLoaded()` |
| `consent_dialog.html` | — | the add-on consent dialog |

`client/common/` holds what every page loads: header, status line, error panel, version banner,
consent modal, the template registry (`SHEET_TEMPLATES`) and base styles.

## What the server injects

Each page's globals script receives the Picker keys, the view (web app or sidebar), the sheet IDs
and sheet type, and two page flags, `getStarted` and `saveFile`. The flags tell the shared consent
code which page to resume. The sheet IDs are reassigned as flows run.

## Calling the server

`runAppsScript(method, …args)` wraps `google.script.run` in a promise and rejects with a normalised
error. A resolved call can still be a failure envelope — check it with `AppError.check`. Bulk work
fans out with `Promise.all`; each item resolves even on failure, so one bad sheet cannot stop the
rest. See [08](08-error-handling.md).

## Consent

Scopes are not granted by opening a page.

- **Web app:** a modal opens Google's consent page in a popup; when the popup closes, the page
  rechecks.
- **Sidebar:** a sidebar cannot host the consent page, so it opens the consent dialog. After the
  user authorizes, the dialog leaves a one-time signal in the user's properties; the sidebar polls
  for it, then rechecks.

On success the consent code resumes the page: `authorizeSaveFile`, `authorizeGetStarted` or
`authorizeAndContinue`, chosen by the page flags.

## Access

Under `drive.file`, knowing a file's ID is not enough. Every workflow checks access, opens the
Google Picker showing only the files that need granting, rechecks and continues. A file the user
neither owns nor edits stops the flow; one they edit but do not own shows a warning.

## Status and mobile

One status line per page: `setStatusWithSpinner` while working, `setStatusText` when done. The
error panel takes over the line when it shows. On mobile, the page swaps in mobile instructions and
shorter status text.

User-controlled names are escaped before they reach the page, and links are allow-listed to Google
Sheets and Drive folders.

## Conventions

- Buttons start hidden and disabled in the section markup; flows reveal them.
- `™` on Google product names; emoji mark status (✅ ❌ ⚠️ 🔐 ⛔ 🎉).
- Commented-out buttons are older flows kept for reference.
