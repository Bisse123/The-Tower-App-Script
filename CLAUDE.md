# The Tower — App Script

Google Apps Script (V8 runtime) project behind the community spreadsheets ("IDS Sheets™") for the
mobile game *The Tower*. One script project with two faces — a Google Sheets add-on and a
standalone web app — serving three workflows: **Get Started**, **Update Sheet** and **Import Data
From Game**. Deployed with clasp 2.x. No build step, no test suite, no linter.

## Structure

```
src/                   everything clasp pushes; a FLAT namespace — see Quirks
  server/core/         errors, app version, cache, Sheets API wrapper
  server/helpers/      versions, label discovery, sheet refs, presets, dropdown values
  server/entry/        sheet-type registry, web-app router, add-on menu, export/import endpoints
  server/workflows/    the client-callable functions, one file per workflow step
  server/savefile/     save-file parser: binary reader, parse orchestration, preferences
  server/sheets/<type>/  one folder per sheet type, files by role: <type>.js (contract),
                       _write, _read, _savefile, _catalog (game ID-to-name tables)
  client/pages/        one page shell per workflow, plus the consent dialog
  client/common/       fragments every page loads: header, status, error panel, consent
  client/<workflow>/   update, get_started, save_file: _section / _styles / *_scripts
documentation/         the human-facing docs, one per area
.claude/             maps and checklists for working in this repo
docs/                save-format reference JSON (git-ignored, local only)
.github/             deploy workflow and its guard scripts
```

## Commands

```bash
npm run sandbox    # push src/ to the dev script project
npm run dev        # push src/ to PRODUCTION HEAD — this is how a change gets tested
npm run bump patch # rewrite the version in src/server/core/app_version.js, then push to production HEAD
npm run check      # the only local verification — run after editing src/
npm run check -- --compare <ref>   # also: every page holds the same code as at <ref>
```

Nothing runs locally and there is no compiler. `npm run check` parses every `.js` file and every
`<script>` block; confirms every `object.member` on a server object is declared, every `include()`
and page name resolves, every server function the client calls exists, and every inline handler
calls a global on its page; flags a global declared twice on one page; loads each page's scripts in
order in a stand-in browser; and loads each server file on its own. After moving client code, run it
with `--compare` against the commit before the move: each page must keep the same lines, globals and
load outcome. `--page old=new` pairs a renamed page.

It does not catch a wrong cell offset or a missing neutral key. Behaviour is only observable after
`npm run sandbox`, which the user runs — so state plainly what was verified and what was not.

## Conventions

Written to an older dialect than the runtime allows: `const` for a module object at file top
level, `var` for locals, `function (…) {}` expressions for members. Match the surrounding file
rather than modernising it.

The full set arrives automatically — `.claude/rules/` holds three path-scoped rules that load when
a matching file is read, so they cost nothing on a task that never touches `src/`:

| Rule | Loads when reading |
| --- | --- |
| `rules/conventions.md` — style, error API, silent-failure rules | any `src/` file |
| `rules/backend.md` — server-side symbol map | `src/*.js` |
| `rules/frontend.md` — client-side symbol map | `src/*.html` |

## Quirks

- **`src/*.js` is one flat global namespace.** Apps Script concatenates every file — there are no
  imports and no modules. Any top-level symbol is visible everywhere, and any top-level `function`
  is callable from the client via `google.script.run`.
- **A literal `//` inside a fragment's `<script>` is stripped as a comment and kills the block.**
  Build URLs from the `googleLink` constant in the page shell instead.
- **A new template version is two edits, not one** — the sheet module's converter *and* the
  matching converter in `server/sheets/ids_collection/`. Only the second one is easy to miss, and it breaks
  single-file users silently.
- **Server functions return failures, they never throw** across the `google.script.run` boundary.
  Return an `errors.fail(...)` envelope; the client switches on its code.

## Documentation

The symbol maps load on their own (see Conventions). These do not:

| I need to… | Read |
| --- | --- |
| Plan a recurring change, or debug by symptom | [.claude/tasks.md](.claude/tasks.md) |
| Ship a release, bump a version, debug CI | [documentation/07-deployment.md](documentation/07-deployment.md) |
| Caching, label discovery, version comparison, file ops | [documentation/01-architecture.md](documentation/01-architecture.md) |
| Template copying and ID cross-linking | [documentation/02-workflow-get-started.md](documentation/02-workflow-get-started.md) |
| The migration flow — largest and most stateful | [documentation/03-workflow-update-sheets.md](documentation/03-workflow-update-sheets.md) |
| The save-file parser or the diff view | [documentation/04-workflow-save-file-import.md](documentation/04-workflow-save-file-import.md) |
| Add a sheet type or a version converter | [documentation/05-sheet-modules.md](documentation/05-sheet-modules.md) |
| Pages, the picker, the consent flow | [documentation/06-frontend.md](documentation/06-frontend.md) |
| Trace a reported failure to a log entry | [documentation/08-error-handling.md](documentation/08-error-handling.md) |
| How a save-file category is encoded | `docs/<category>_save_format.json` |
