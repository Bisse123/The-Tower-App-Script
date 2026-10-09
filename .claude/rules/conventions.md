---
paths:
  - "src/**/*.{js,html}"
---

# Conventions

## Rules that fail silently

- **No load-time cross-file use.** A file may reference another file's symbols only inside a
  function. Each new file declares its own object rather than adding members to another file's.
- **Never delete a version converter.** Old sheets still need their reader.
- **Converters return the same neutral keys.** `importData` branches on key presence; a renamed key
  imports nothing and reports success.
- **`getVersionXX*` readers never call the API.** Raw values in, plain objects out.
- **A new template version is two edits:** the sheet type's converter and the IDS Collection's.
- **Never hard-code a cell address.** Find a text label and read at an offset from it.
- **No literal `//` in an included fragment's `<script>`.** Build URLs from `googleLink`.
- **In `getTemplateAndsheetIds`, a sheet type whose name is a substring of another comes after it.**
- **Page order is load-bearing:** styles, sections, the page's globals `<script>`, then fragment
  scripts. A function is hoisted only within its own `<script>` block, so code that runs on load may
  only call earlier blocks.
- **`include()` inlines verbatim; `includeTemplate()` evaluates scriptlets.** Only `error_scripts`
  needs the template form.

## Style

- Apps Script V8 in an older dialect: `const` module objects, `var` locals, `function (…) {}`
  members. Match the file.
- Module methods open with `console.log("Called: <object>.<method>")`.
- JSDoc above every function: what it does, each parameter, the return value.
- No comments that narrate a change or justify an edit.
- `™` on user-visible Google product names ("Google Sheet™").
- Emoji status vocabulary: ✅ success · ❌ failure · ⚠️ partial · 🔐 access · ⛔ blocked · 🎉 done.
- Commented-out UI is kept on purpose; the live element is the uncommented one.
- `server/sheets/ids_collection/*.js` carry a UTF-8 BOM — keep it. Never edit `.clasp.json` or
  `remote_head/`.

## Errors — server

Nothing throws across `google.script.run`; failures return an envelope with a code.

| Situation | Call |
| --- | --- |
| Any `catch` | `errors.report(source, error, context)` → `errors.fail(report)` |
| A precondition you checked | `errors.reject(source, code, message)` |
| An inner call already failed | `errors.propagate(source, inner, message?)` — never `reject` |
| A `catch` that logs and carries on | `errors.report(source, error, context, errors.CODES.RECOVERED)` |
| A `catch` that is one of the expected answers | pass the code, e.g. `errors.CODES.ACCESS_DENIED` |

- `source` is `functionName`, or `object.method` for an object member.
- `context` takes the function's parameters, raw; `errors.snapshot` bounds it. Never a raw email.
- **Code choice:** would *we* have to change something? Bug (`INVALID_INPUT`, `SHEET_STRUCTURE`,
  `CLIENT`, `INTERNAL`). Otherwise expected (`ACCESS_DENIED`, `NOT_FOUND`, `INVALID_LINK`,
  `INVALID_FILE`, `QUOTA`, `TIMEOUT`, `VERSION_OUTDATED`, `RECOVERED`, `AUTH_UNAVAILABLE`,
  `NETWORK_BLOCKED`).
- New codes go in `ERROR_DEFS` only (plus a title in `AppError.TITLES`).
- Never tell the user to update their sheet via `MESSAGES.VERSION_OUTDATED`; the call site says it.
- Never `console.log` an error.

## Errors — client

Call the server with `runAppsScript(method, …args)`, a promise over `google.script.run`. A resolved
call can still be a failure envelope.

| Situation | Call |
| --- | --- |
| A failed envelope | `AppError.check(result, source)` / `AppError.show(result, { source })` |
| A caught exception | `AppError.show(error, { source, message })` |
| A failure the user need not see | `AppError.log(error, source)` |
| A list of per-sheet failures | `AppError.surfaceBatch(entries, { source })` — each entry carries its `envelope` |

Fan-outs: `resolve` on failure, never `reject`; update the DOM inside the handler.

Escape user-controlled names with `escSaveFileHtml` / `escapeSummaryHtml`; pass URLs through
`sanitizeGetStartedUrl` before rendering a link.
