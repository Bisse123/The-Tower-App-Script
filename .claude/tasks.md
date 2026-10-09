# Task checklists

Recurring changes span several files, and a missed step usually fails silently. Run
`npm run check` after each.

## A new template version of an existing sheet type

1. In `server/sheets/<type>/<type>_read.js`, add `versionN_M(oldSheetID)` and its pure
   `getVersionN_M*(values)` readers.
2. Register `"vN.M"` in `convertVersionFunctions` in `<type>.js`. Order does not matter.
3. Update `importData` and the `_write.js` functions for the new layout.
4. Mirror it in `server/sheets/ids_collection/`: a converter and the import branch.
5. If the template file changed, update `SHEET_TEMPLATES` in `client/common/templates_scripts.html`.

## A new sheet type

1. Add `server/sheets/<type>/` with the contract file and the role files it needs.
2. Register it in `sheetVars()` (`server/entry/registry.js`).
3. Add it to `getTemplateAndsheetIds` (`server/workflows/templates.js`) after any type whose name
   contains it, and to the default list in `getSaveFileImportTargets`.
4. Add it to `SHEET_TEMPLATES`, `showSelectImportSection()` and the conversion type list in
   `importData()` (`client/update/transfer_scripts.html`).
5. Save-file support: a field map and `parse*Data` in `_savefile.js`, its call in
   `server/savefile/parse.js`, and renderers in `client/save_file/diff_scripts.html` and
   `body_scripts.html`.
6. Add it to the IDS Collection if it belongs in the single-file arrangement.

The template needs a `Home Page` with version labels and a copy link, and an `IDS` tab with its own
ID and the IDS Master's.

## A new field in the save file

1. Find the game's key in `docs/<category>_save_format.json`.
2. Add it to the type's field map in `_savefile.js` and read it in `parse*Data`.
3. If `importData` does not write it yet, extend `_write.js` and the IDS Collection's import.
4. Show it in the diff and body renderers.

`parse*Data` must tolerate the field being absent — older save files lack it.

## Debugging by symptom

| Symptom | Start at |
| --- | --- |
| "Could not find sheet ID for X" | `labelUtils.findSheetTypeID` / `findSheetTypeURL` — a template label moved |
| Stale data after an import | `CacheManager.RemoveSpreadsheet` — a mutation that did not invalidate |
| A picker keeps asking for access | The access cycle; which bucket the file landed in |
| A dropdown rejects the written value | `dropdownUtils.getDVTValue` and the named-range tables in `importData` |
| Presets land in the wrong slots | `presetUtils.resolvePresetOrder` |
| A category imports nothing but reports success | A converter renamed a neutral key |
| A fragment's functions are "not defined" | A literal `//` in it was stripped, or a load-order problem: `npm run check` |
| `ReferenceError: getStarted` on the save-file page | The consent `saveFile` branch must run first |

A user quoting a `TWR-…` reference:
[08 ▸ Runbook](../documentation/08-error-handling.md#runbook). CI and deploy failures:
[07 ▸ Troubleshooting](../documentation/07-deployment.md#troubleshooting).
