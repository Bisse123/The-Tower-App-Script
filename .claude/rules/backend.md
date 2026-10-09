---
paths:
  - "src/**/*.js"
---

# Backend map

All `src/**/*.js` share one global namespace. Search for a symbol rather than trusting a path.

## Layers

Calls run downward only: client → `entry`, `workflows`, `savefile` → `sheets` → `helpers` →
`core` → Sheets v4 · Drive v3 · CacheService · PropertiesService.

## server/core

| File | Holds |
| --- | --- |
| `errors.js` | `ERROR_DEFS` (every code) and `errors`: `report` · `fail` · `reject` · `propagate` · `snapshot` |
| `error_endpoints.js` | Client-callable `reportClientError` · `reportServerError` · `errorContract` |
| `app_version.js` | `appVersion` (`VERSION`/`MINIMUM` rewritten by `npm run bump`), `getAppVersionStatus`, `setLatestAppVersion` (CI) |
| `cache.js` | `CacheManager` — per-user, 60 s, chunked; spreadsheets keyed by logical name |
| `sheets_api.js` | `SheetsAPI` — metadata, cached `batchGetValues`/`batchGetFormulas`, `batchUpdateValues`, `applySheetVisibility` |

## server/helpers

| Object | Holds |
| --- | --- |
| `versionUtils` | `findSheetVersion` · `getEPathsVersion` · `isVersionLoading` · `readVersion` · `getVersionStatus` · `compareVersions` |
| `labelUtils` | `isSheetTypeCell` · `findSheetTypeID` · `findSheetTypeURL` · `findSheetTemplateID` · `addIDUpdatesToBatch` |
| `sheetRefs` | `isSheetId` · `extractSheetId` · `columnToLetter` · `extractUrlFromHyperlink` · `getColumnOffsetFromRange` |
| `presetUtils` | `templatePresetNames` · `resolvePresetOrder` · `presetUnlocks` · `unlockedPresetCounts` |
| `dropdownUtils` | `getDVTValue` — maps a level onto a value the dropdown accepts |

## server/entry

| File | Holds |
| --- | --- |
| `registry.js` | `sheetVars(sheetType)` — sheet type → contract object (add new types here); `spreadsheets(name, id)` |
| `web_app.js` | `doGet` router (`?page=savefile` · `?page=getstarted` · update page), `include`, `includeTemplate` |
| `addon.js` | Menu, the three dialogs, consent signal, sidebar `get*Parameters` |
| `transfer.js` | `exportData` · `importData` — the dispatchers every workflow uses |

## server/workflows

`access.js` (token, scopes, sheet/template access) · `templates.js` (`getTemplateAndsheetIds`, holding
the candidate sheet-type list) · `version_checks.js` (`compareSheetVersions`,
`checkExportCompatibility`) · `update_prepare.js` (`prepareImportData`) · `file_moves.js` (copy,
move, convert-move, delete) · `ids_master.js` (`updateIdsMaster`, `getIdsMasterGid`,
`updateSheetID`) · `get_started.js` · `save_file_targets.js`.

## server/savefile

`parse.js` (`parseSaveFileBytes`: ungzip → `parseNRBF` → each type's field map and parser) ·
`nrbf_reader.js` (.NET BinaryFormatter reader) · `preferences.js` (wave-cap preference).

## server/sheets/`<type>`

One folder per sheet type, one object per role (Cards as the example):

| File | Object | Holds |
| --- | --- | --- |
| `<type>.js` | `cards` | Contract: `exportData`, `importData`, `convertVersionFunctions`, `isCompatibleVersion` |
| `_write.js` | `cardsWriter` | `update*` — build the batch writes; also called by the IDS Collection |
| `_read.js` | `cardsReader` | `versionN_M` converters and pure `getVersionN_M*` readers |
| `_savefile.js` | `cardsSaveFile` | The save-file field map and `parse*Data` |
| `_catalog.js` | `cardsCatalog` | Game ID-to-name tables (only where a type has them) |

`exportData` looks up the converter from `isCompatibleVersion`'s key; export and save-file parsing
both return the same neutral object `importData` takes. `importData` guards every section and
reads the new sheet before writing it.

Variants: Modules splits its readers per tab and has `modulesLayout`; Effective Paths keeps each
tab's writer and readers together (`ePathsEHP`, `ePathsEDamage`, `ePathsEEcon`); IDS Collection
splits its converters by template generation (`collectionConvertersV4` …); `legacy/` holds the old
Relics and Themes & Songs whole.

**IDS Master** moves the `IDS` registry and preset names; its import writes its own IDs.
**IDS Collection** is every category as tabs in one file: its export is keyed by sheet type, and
its import calls each type's writers and collects failures in `failedUpdates`.
