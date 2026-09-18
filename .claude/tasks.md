# Task checklists

The file fan-out for each recurring change. Most span several layers, and a missed step usually
fails silently rather than loudly. The rules behind these are in
[conventions.md](rules/conventions.md).

---

## A new template version of an existing sheet was released

1. Add `versionN_M: function (oldSheetID) { … }` to the sheet module, fetching what the new
   template needs.
2. Add its `getVersionN_M*(values)` pure readers — no API calls.
3. Register the threshold in `get convertVersionFunctions()`:
   `"vN.M": this.versionN_M.bind(this)`. Declaration order does not matter; `isCompatibleVersion`
   sorts by version.
4. Update the module's `importData` and its `update*` helpers for the new layout.
5. **Mirror the change in `src/14_IDS_Collection.js`** — its own converter chain plus the branch
   calling the module's `update*` against the collection's tabs.
6. If the template file ID changed, update `SHEET_TEMPLATES` in `src/21_templates_scripts.html`.

---

## A brand-new sheet type must be supported

1. Create `src/NN_<Name>.js` with a `const <object> = { … }` implementing the full contract —
   see [backend.md ▸ The contract](rules/backend.md#the-contract).
2. Register it in `sheetVars()` in `src/01_Main.js`.
3. Add it to the candidate list in `getTemplateAndsheetIds` (`src/02_Shared.js`), respecting the
   substring-ordering rule.
4. Add it to the client type lists:
   - `showSelectImportSection()` in `src/24_selectImport_scripts.html`
   - `SHEET_TEMPLATES` in `src/21_templates_scripts.html`
   - the sheet-type list in `src/25_fileAccess_scripts.html`
   - the default list in `getSaveFileImportTargets` (`src/02_Shared.js`)
5. For save-file support: a header map and a `parse*Data` call in `src/02_SavedFile.js`, and a
   diff renderer branch in `src/28_saveFile_scripts.html`.
6. Add the category to `src/14_IDS_Collection.js` if it belongs in the single-file arrangement.

The template itself must provide a `Home Page` carrying version labels and a copy link, and an
`IDS` tab carrying its own ID and the IDS Master's.

---

## The game added a new field to the save file

1. Confirm the game's key name against `docs/<category>_save_format.json`.
   `module_save_format.json` carries the complete substat `effectID` encoding table.
2. Add `ourName: "gameKey"` to the category's header map in `src/02_SavedFile.js`.
3. Read `data.ourName` inside the module's `parse*Data` and emit it under the key `importData`
   already expects.
4. If `importData` does not yet write it, extend the module's `update*` **and** the corresponding
   branch in `src/14_IDS_Collection.js`.
5. Add a diff renderer branch in `src/28_saveFile_scripts.html` so the change is visible before
   import.

Every `parse*Data` guards with `hasOwnProperty` / null checks — an older save file missing a field
must parse cleanly, just without that value.

---

## Debugging by symptom

| Symptom | Start at |
| --- | --- |
| "Sheet not found" / "Could not find sheet ID for X" | `shared.findSheetTypeID` / `findSheetTypeURL` — a label moved in the template |
| Stale data after an import | `CacheManager.RemoveSpreadsheet` — a mutation that did not invalidate |
| A picker keeps re-asking for access | The access-grant cycle; check which bucket the file landed in |
| A dropdown cell rejects the written value | `shared.getDVTValue` and the named-range tables in `14_IDS_Collection.js` |
| Presets land in the wrong slots | `shared.resolvePresetOrder` and `shared.templatePresetNames` |
| A category imports nothing but reports success | A converter renamed a neutral key, or `importData`'s guard never matched |
| A fragment's script block does nothing | A literal `//` inside it was stripped as a comment |
| Version comparison odd on Effective Paths | Four-segment zero-padded versions; it uses `shared.getEPathsVersion` |
| A `ReferenceError` on `getStarted` in the save-file page | `20_SavedFileApp.html` never declares it — the `saveFile` branch must short-circuit first |

A user quoting a `TWR-…` reference, or an empty Error Reporting console:
[08-error-handling.md ▸ Runbook](../documentation/08-error-handling.md#runbook).
CI, deploy and picker-initialisation failures:
[07-deployment.md ▸ Troubleshooting](../documentation/07-deployment.md#troubleshooting).
