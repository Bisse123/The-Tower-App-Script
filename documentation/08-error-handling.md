# 08 — Error handling

One path from a failure to a Cloud Logging entry and a message the user can act on.

| | |
| --- | --- |
| Server | `server/core/errors.js` (`errors`, `ERROR_DEFS`), `server/core/error_endpoints.js` |
| Client | `client/common/error_*` (`AppError`, `runAppsScript`, the panel) — on every page |

## Two kinds of failure

The code decides everything.

| | Expected | Bug |
| --- | --- | --- |
| Means | The app working as designed on input it cannot accept | Something we got wrong |
| Severity | `WARNING` | `ERROR`, sent to Error Reporting |
| Panel | amber, no reference | red, with a `TWR-…` reference to copy |

| Code | Kind | When |
| --- | --- | --- |
| `ACCESS_DENIED` | expected | Drive or Sheets refused, or a file was never granted |
| `NOT_FOUND` | expected | The file is gone or was never shared |
| `INVALID_LINK` | expected | The input is not a sheet link or ID |
| `INVALID_FILE` | expected | The picked file is not a valid `playerInfo.dat` |
| `QUOTA` | expected | Google is rate-limiting the account |
| `TIMEOUT` | expected | Execution time ran out |
| `VERSION_OUTDATED` | expected | The sheet is too old for this step |
| `AUTH_UNAVAILABLE` | expected | Sign-in cannot complete in the browser |
| `NETWORK_BLOCKED` | expected | A request never reached Google |
| `RECOVERED` | expected | Logged, and the script carried on — never shown |
| `INVALID_INPUT` | bug | A required parameter never arrived |
| `SHEET_STRUCTURE` | bug | A tab or label the code looks for is missing |
| `CLIENT` | bug | A failure in the page |
| `INTERNAL` | bug | Anything unclassified |

Every code is declared once in `ERROR_DEFS`; the page receives the same table. `errors.classify`
sorts exceptions by Google's wording (a bad range is `SHEET_STRUCTURE`, a 429 is `QUOTA`). The page
recognises sign-in and network failures by their text, and ignores failures raised by browser
extensions.

## The envelope

Server functions return this on failure and never throw:

```javascript
{ success: false, code, expected, message, reference, detail, trace, outdated? }
```

`message` is for the user; `detail` and `trace` belong to the layer that failed and show only under
*Technical details*. `outdated` appears when the running version is behind (see
[07](07-deployment.md#versioning)).

## How a failure travels

- **Caught exception:** `errors.report` classifies it; `errors.fail` builds the envelope; each layer
  above uses `errors.propagate`, which extends `trace`. An expected failure is logged immediately;
  a bug is logged when the page shows it and calls `reportServerError` — so a bug in a tab that
  closes first is never logged.
- **Precondition the code checked:** `errors.reject`, same split.
- **Recovered:** `errors.report(…, RECOVERED)` logs a warning and returns nothing.
- **Browser failure:** `AppError` sends it to `reportClientError`, logged as
  `the-tower-app-script-client`.
- Identical failures from one user within 5 minutes become one entry; later ones reuse its
  reference.

Client calls: `AppError.show` (panel) · `AppError.check` (show if failed) · `AppError.log` (record
silently) · `AppError.surfaceBatch` (per-sheet lists: panel for the first bug).

## Google Cloud

Each entry carries `reference`, `code`, `kind` (`expected` or `bug`), `source` (the deepest frame),
`trace`, `detail`, optional `note` and `data`, the app version and a hashed user key. Error Reporting
groups by the stack trace's first line, so `errors.text` and `errors.stack` must keep it non-empty.

Setup per GCP project:

1. Enable the Error Reporting API.
2. Grant on-call *Error Reporting Viewer* and *Logs Viewer*.
3. Add a counter metric `app_script_errors` on `severity>=ERROR AND jsonPayload.code!=""`, labelled
   by `code`, `source` and `kind`, and an alert on it.

Never log a raw email, password, token or payment detail. For a bug, `note`, `data` and `detail`
pass through the browser on the way to the log.

## Runbook

| Find | Query |
| --- | --- |
| A quoted reference | `jsonPayload.reference="TWR-M4X2K9-A7F3"` |
| Bugs only | `jsonPayload.kind="bug"` |
| One failing function | `severity>=ERROR AND jsonPayload.source="collection.importData"`, grouped by `jsonPayload.context.user` |
| One release | `severity>=ERROR AND jsonPayload.serviceContext.version="<version>"` |
| Browser failures | `jsonPayload.serviceContext.service="the-tower-app-script-client"` |

No entry? The bug's round trip never completed, or it was folded into an identical entry within 5
minutes.
