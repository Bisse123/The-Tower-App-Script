---
paths:
  - "src/*.js"
  - "src/*.html"
  - "src/**/*.{js,html}"
---

# Conventions

How to write code in this repo. The distilled form of what
[08-error-handling.md](../../documentation/08-error-handling.md) and the frontend doc spell out at
length — enough to write correct code, not enough to debug a live incident.

---

## Rules that fail silently

These produce no exception, no log entry and no failed check. They surface as wrong data in a
user's spreadsheet.

- **Never delete an old version converter.** A user on the oldest supported template still needs
  its reader. Converters are added, never removed.
- **A converter must return the same neutral keys.** `importData` branches on key presence, so a
  renamed key imports nothing *and reports success*.
- **`getVersionXX*` readers must not call the API.** Raw values in, plain objects out.
- **A new template version is at least two edits** — the sheet module's converter *and* the
  matching branch in `14_IDS_Collection.js`. Missing the second breaks single-file users only.
- **Never hard-code a cell address in a user's sheet.** Scan for a text label and read at a fixed
  offset from it.
- **A literal `//` inside a fragment's `<script>` is stripped as a comment and kills the block.**
  URLs belong in the page shell as a constant (`googleLink`, no trailing slash), never in a
  fragment.
- **In `getTemplateAndsheetIds`, a sheet type whose name is a substring of another must be looked
  up after it.** `Themes, Songs & Relics` precedes `Relics`.
- **Fragment order in a page shell is load-bearing** — styles in `<head>` → sections → the
  server-injected globals `<script>` → fragment scripts, which run at parse time and read those
  globals immediately.
- **`include()` inlines verbatim; `includeTemplate()` evaluates scriptlets.** A stray `<?` in an
  `include()`d fragment is a parse error. Only `22_error_scripts` needs the template form.

---

## Style

Apps Script V8, written to an older dialect throughout. Match what is already there.

- `const <name> = { … }` for a module object at file top level; `var` for locals inside functions;
  `function (…) {}` expressions for members — not arrow functions or method shorthand.
- Every module method opens with `console.log("Called: <object>.<method>")`.
- JSDoc above every function, method and property: what it is, each parameter, what it returns.
- Never write a comment that narrates a change, justifies an edit, or compares to previous
  behaviour. Remove such comments from any code being edited.
- `™` on user-visible Google product names — `"Google Sheet™"`, `"New spreadsheet™ not found"`.
- Emoji as status vocabulary: ✅ success · ❌ failure · ⚠️ partial · 🔐 access needed · ⛔ blocked ·
  📂 files · 🔄 update · 🎉 all done.
- Commented-out UI is left in place deliberately; the live element is the one without comment
  markers.
- Prefer `Edit` over rewriting a file with `Write`. `14_IDS_Collection.js` carries a UTF-8 BOM —
  leave it. Never hand-edit `.clasp.json` (transient) or anything under `remote_head/`.

---

## Errors — server

**No server function may throw across the `google.script.run` boundary.** Every one returns a
success flag; a failure carries a code the client switches on.

| Situation | Call |
| --- | --- |
| `catch` around anything | `errors.report(source, error, context)` then `errors.fail(report)` |
| A precondition you checked yourself | `errors.reject(source, code, message)` |
| An inner call already failed | `errors.propagate(source, inner, message?)` — never `reject`, or one incident is recorded twice |
| A `catch` that logs and carries on | `errors.report(source, error, context, errors.CODES.RECOVERED)` |
| A `catch` that is one of the answers the function was called to give | Pass the code explicitly, e.g. `errors.CODES.ACCESS_DENIED` |

`source` is `functionName` for a top-level function and `module.method` for a sheet-module method
— fifteen modules share those method names, so the qualifier is what identifies the failure.

`context` takes the function's own parameters, raw — `errors.snapshot` caps depth and size, so
raw locals are safe to pass. Add a mid-computation local when it would narrow down the failure.
Never pass a raw email; `errors.userKey()` already identifies the user as a truncated hash.

Never `console.log` an error: INFO severity, no stack, nothing will ever alert on it.

### Choosing a code

**Would *we* have to change something?** Then it is a bug — severity `ERROR`, a `TWR-…` reference
shown to the user, Error Reporting sees it. Otherwise expected — `WARNING`, no reference, amber
panel, and it never reaches Error Reporting.

Expected: `ACCESS_DENIED` · `NOT_FOUND` · `INVALID_LINK` · `INVALID_FILE` · `QUOTA` · `TIMEOUT` ·
`VERSION_OUTDATED` · `RECOVERED` · `AUTH_UNAVAILABLE` · `NETWORK_BLOCKED`.
Bugs: `INVALID_INPUT` · `SHEET_STRUCTURE` · `CLIENT` · `INTERNAL`.

A new code goes in `ERROR_DEFS` (`00_Errors.js`) and nowhere else, bar the display title in
`AppError.TITLES`.

**Never tell the user to update their sheet** via `MESSAGES.VERSION_OUTDATED` — it says only that
the sheet is not a version the step can work with. Every call site that knows more says it itself.

---

## Errors — client

`runAppsScript(method, …args)` is the only way to call the server. A resolved call can still be a
failure — that is what the envelope is for:

```javascript
const result = await runAppsScript("importData", newSheetID, sheetType, data);
if (AppError.check(result, "importData")) return;   // panel shown, reference included
```

| Situation | Call |
| --- | --- |
| A failed envelope | `AppError.check(result, source)` / `AppError.show(result, { source })` |
| A caught exception | `AppError.show(error, { source, message })` |
| A failure the user need not see | `AppError.log(error, source)` |
| A list of per-sheet failures | `AppError.surfaceBatch(entries, { source })` |
| Several failures at once | `AppError.showAll(rawList, { source })` |

Each entry in a batch must carry its `envelope`; without it there is no code and no reference.

### Fan-outs

- **`resolve` on failure, never `reject`** inside a `Promise.all`, or one bad sheet takes down the
  whole batch.
- **Update the DOM inside the handler**, not after the `await`.
- **Report what you swallowed** — `AppError.log(error, "copyTemplates")` records a
  resolved-on-failure item without showing the panel.

### Escaping

Sheet and file names are user-controlled and reach `innerHTML`. Use `escSaveFileHtml` /
`escapeSummaryHtml` for names, and `sanitizeGetStartedUrl` before rendering any URL into a link —
it allow-lists only `docs.google.com/spreadsheets/d/…` and `drive.google.com/drive/folders/…`.
