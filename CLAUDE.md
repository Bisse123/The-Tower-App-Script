# The Tower — App Script

Google Apps Script (V8) project behind the community spreadsheets ("IDS Sheets™") for the mobile
game *The Tower*. One script project serves a Sheets add-on and a web app, with three workflows:
**Get Started**, **Update Sheet** and **Import Data From Game**. Deployed with clasp 2.x. No build,
tests or linter.

## Structure

```
src/                     everything clasp pushes — one flat global namespace
  server/core/           errors, app version, cache, Sheets API wrapper
  server/helpers/        versions, label discovery, sheet refs, presets, dropdown values
  server/entry/          sheet-type registry, web-app router, add-on menu, export/import
  server/workflows/      client-callable functions, one file per workflow step
  server/savefile/       save-file parser
  server/sheets/<type>/  one folder per sheet type: <type>.js (contract), _write, _read,
                         _savefile, _catalog (game ID-to-name tables)
  client/pages/          page shells
  client/common/         fragments every page loads
  client/<workflow>/     update, get_started, save_file
documentation/           human docs
.claude/                 maps and checklists
docs/                    save-format reference JSON (git-ignored)
```

## Commands

```bash
npm run sandbox                   # push src/ to the dev project
npm run dev                       # push src/ to PRODUCTION HEAD
npm run bump patch                # bump the version, then npm run dev
npm run check                     # the only local verification — run after editing src/
npm run check -- --compare <ref>  # also: every page holds the same code as at <ref>
```

`npm run check` parses every file; resolves every server `object.member`, include, page name,
client-to-server call and inline handler; flags duplicate page globals; and loads each page's
scripts in order and each server file alone. Use `--compare` after moving client code
(`--page old=new` pairs a renamed page). It cannot catch a wrong cell offset or a missing neutral
key: behaviour shows only after `npm run sandbox`, which the user runs. Say what was verified and
what was not.

## Conventions

Older dialect: `const` module objects, `var` locals, `function (…) {}` members. Match the file.
`.claude/rules/` loads automatically: `conventions.md` for any `src/` file, `backend.md` for
`.js`, `frontend.md` for `.html`.

## Quirks

- **No file may use another file's symbols while loading.** Push order is not guaranteed.
- **Every top-level server `function` is client-callable** by name; renaming one breaks callers.
- **A literal `//` in an included fragment's `<script>` is stripped as a comment** and kills the
  block. Build URLs from `googleLink` in the page shell.
- **A new template version is two edits:** the sheet type's converter and the IDS Collection's.
- **Server functions return failures, never throw** across `google.script.run`.

## Documentation

| Need | Read |
| --- | --- |
| A recurring change, or debugging by symptom | [.claude/tasks.md](.claude/tasks.md) |
| Release, versioning, CI | [documentation/07-deployment.md](documentation/07-deployment.md) |
| Cache, label discovery, versions, file moves | [documentation/01-architecture.md](documentation/01-architecture.md) |
| A workflow end to end | [02](documentation/02-workflow-get-started.md) · [03](documentation/03-workflow-update-sheets.md) · [04](documentation/04-workflow-save-file-import.md) |
| Sheet types and converters | [documentation/05-sheet-modules.md](documentation/05-sheet-modules.md) |
| Pages, picker, consent | [documentation/06-frontend.md](documentation/06-frontend.md) |
| A reported failure | [documentation/08-error-handling.md](documentation/08-error-handling.md) |
| How the save file encodes a category | `docs/<category>_save_format.json` |
