# The Tower — App Script

Google Apps Script (V8 runtime) project behind the community spreadsheets ("IDS Sheets™") for the
mobile game *The Tower*. One script project with two faces — a Google Sheets add-on and a
standalone web app — serving three workflows: **Get Started**, **Update Sheet** and **Import Data
From Game**. Deployed with clasp 2.x. No build step, no test suite, no linter.

## Structure

```
src/             everything clasp pushes; a FLAT namespace — see Quirks
  00_*.js        errors, versioning
  01–02_*.js     entry points, shared infrastructure, save-file parser
  03–17_*.js     one module per sheet type, all to the same contract
  20_*.html      one page shell per workflow
  21–29_*.html   UI fragments in triples: _section / _styles / _scripts
documentation/   the human-facing docs, one per area
.claude/         maps and checklists for working in this repo
docs/            save-format reference JSON (git-ignored, local only)
.github/         deploy workflow and its guard scripts
```

## Commands

```bash
npm run sandbox    # push src/ to the dev script project
npm run dev        # push src/ to PRODUCTION HEAD — this is how a change gets tested
npm run bump patch # rewrite the version in src/00_Version.js, then push to production HEAD
```

Nothing runs locally. There is no compiler, so these two syntax checks are the only verification
available — run both after editing `src/`:

```bash
for f in src/*.js; do node --check "$f" || echo "FAILED: $f"; done
```

```bash
node -e '
const fs=require("fs");
for(const f of fs.readdirSync("src").filter(f=>f.endsWith(".html"))){
  for(const b of (fs.readFileSync("src/"+f,"utf8").match(/<script>([\s\S]*?)<\/script>/g)||[])){
    const code=b.replace(/^<script>/,"").replace(/<\/script>$/,"");
    if(code.includes("<?")) continue;
    try{ new Function(code); }catch(e){ console.log("SYNTAX "+f+": "+e.message); }
  }
}'
```

Neither catches a wrong cell offset or a missing neutral key. Behaviour is only observable after
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
  matching branch in `14_IDS_Collection.js`. Only the second one is easy to miss, and it breaks
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
