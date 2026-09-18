---
paths:
  - "src/*.js"
  - "src/**/*.js"
---

# Backend map

Where every server-side symbol lives. Apps Script concatenates all of `src/*.js` into one global
namespace, so any `const` or `function` at file top level is visible everywhere, and any top-level
`function` is callable from the client via `google.script.run`.

Names here are greppable — search for the symbol rather than trusting a line number.

---

## Layering

```
client (google.script.run)
  ├─→ 01_Main.js          entry points, dispatchers
  ├─→ 02_Shared.js        orchestration functions (top level)
  ├─→ 02_SavedFile.js     parseSaveFileBytes, wave-cap preference
  └─→ 00_Version.js       getAppVersionStatus, setLatestAppVersion
        ↓
      02_Shared.js objects: CacheManager · SheetsAPI · shared
        ↓
      03–17 sheet modules
        ↓
      00_Errors.js (every layer reports through it)
        ↓
      Sheets v4 · Drive v3 · CacheService · PropertiesService
```

The client calls Layer 1 **and** Layer 2 directly. There is no single façade.

---

## 00_Errors.js — `errors`

Declarations: `ERROR_REPORT_TYPE`, `CLIENT_ERROR_THROTTLE_SECONDS`, `SNAPSHOT_MAX_DEPTH`,
`SNAPSHOT_MAX_NODES`, `SNAPSHOT_MAX_CHARS`, `ERROR_DEFS`, `errorTable`, `errors`.

`ERROR_DEFS` is the single declaration of every code as `{ expected, client, message }`.
`errors.CODES`, `errors.MESSAGES` and `errors.EXPECTED` derive from it, and so does the client's
copy via `errorContract()`.

Public: `report` · `fail` · `reject` · `propagate` · `snapshot`.
Supporting: `isExpected` · `known` · `contract` · `merge` · `budget` · `reference` · `classify` ·
`text` · `stack` · `userKey` · `record` · `_event` · `_write` · `_runningVersion` ·
`_attachOutdated`.

Top-level, client-callable: `reportClientError(payload)` · `reportServerError(payload)` ·
`errorContract()`. Module-private helpers: `_boundInboundData` · `_throttleReference`.

The error **code** decides everything — expected vs. bug, severity, whether a reference is shown,
whether Error Reporting sees it. Nothing else is consulted.

---

## 00_Version.js — `appVersion`

`VERSION` and `MINIMUM` are string constants rewritten by `npm run bump`; never edit them by hand.
Members: `running` · `latest` · `minimum` · `isOlder` · `status` · `publish`.

Top-level: `getAppVersionStatus()` (client-callable, drives the outdated banner) and
`setLatestAppVersion()` (called by CI through the Apps Script API after a redeploy).

Both are top-level because `google.script.run` and the Apps Script API can only reach top-level
functions, not object members.

---

## 01_Main.js — entry points

| Symbol | Role |
| --- | --- |
| `sheetVars(sheetType)` | **The sheet-type registry.** Maps every sheet-type string to its module object. A new sheet type must be added here. |
| `spreadsheets(typeName, sheetID)` | Cached spreadsheet metadata by logical name, e.g. `"Laboratory newSpreadsheet"`. |
| `doGet(e)` | Web-app router: `?page=savefile` · `?page=getstarted` · default update page. |
| `include` / `includeTemplate` | The fragment inlining used by every page shell. |
| `onOpen` / `onInstall` / `createMenu` | Add-on menu `Import Data`. |
| `showGetStartedDialog` · `showUpdateDialog` · `openSaveFileDialog` | The three add-on views (modal, sidebar, modal). |
| `showAddonConsentDialog` · `markAddonConsentReadySignal` · `consumeAddonConsentReadySignal` | Sidebar consent handshake via `UserProperties`. |
| `getGetStartedParameters` · `getUpdateDialogParameters` · `getSaveFileParameters` | Sidebar context bootstrap, read from the active spreadsheet. |
| `findIdMasterIdInIdsTab(idsSheet)` | Scans an `IDS` tab for the IDS Master's ID. |
| `exportData(oldSheetID, sheetType, versionDifference)` | Generic dispatcher → `module.exportData`. Also captures tab visibility. |
| `importData(newSheetID, sheetType, data, sheetVisibility, idMasterID)` | Generic dispatcher → `module.importData`. Applies visibility first. |

`exportData` and `importData` are the single pair both the migration workflow and the save-file
workflow go through.

---

## 02_Shared.js

Three objects, then a run of top-level orchestration functions.

### `CacheManager`

Per-user cache in front of every Sheets and Drive read. `CacheService.getUserCache()`, 60-second
TTL, 90 000-byte chunk threshold (`CHUNK_SIZE`).

`getSpreadsheet` · `getSheetValues` · `getSheetFormulas` · `getFile` · `RemoveSpreadsheet`.
Internals: `_byteLength` · `_chunkString` · `_retrieveValue` · `_entryKeys`.

The spreadsheet cache is keyed by **logical name**, not ID. A request with a different ID
invalidates and refetches; a request with *no* ID returns whatever is cached. A partial chunked
read is treated as a miss, never as a partial value.

`RemoveSpreadsheet` is called after `moveSheet`, `moveConvertedSheet` and `updateIdsMaster` — the
three places a file's identity or contents change under a cached copy.

### `SheetsAPI`

`fetchSpreadsheet` · `getSheetByName` · `getSheetBySubstring` · `batchGetValues` ·
`batchGetFormulas` · `batchUpdateValues` · `applySheetVisibility`.

Import functions never write incrementally: they accumulate a `batchUpdate` array of
`{ range, values }` across every sub-update and flush once with `SheetsAPI.batchUpdateValues`.

### `shared`

Version reading and comparison:
`findSheetVersion` · `getEPathsVersion` · `isVersionLoading` · `readVersion` · `getVersionStatus` ·
`compareVersions`.

Discovery by label scanning — the app never hard-codes a cell address in a user sheet:
`isSheetTypeCell` · `findSheetTypeID` · `findSheetTypeURL` · `findSheetTemplateID`.

ID and range helpers:
`isSheetId` (44 chars of `[A-Za-z0-9_-]`) · `extractSheetId` · `columnToLetter` ·
`getColumnOffsetFromRange` · `extractUrlFromHyperlink`.

Writing:
`getDVTValue` (maps a `"12 | something"` level string onto a Data-Validation-Table named range, so
dropdown cells receive a value the dropdown accepts) · `addIDUpdatesToBatch` (the two writes every
import performs: `This Sheet ID` and `IDS Master's`).

Presets:
`templatePresetNames` (`["Farming", "Tourney"]`) · `resolvePresetOrder`, which pulls literally-named
presets into the two fixed slots wherever they appear in the save file and returns both the
slot-ordered names and the `indices`, so parallel arrays reorder identically.

### Top-level functions

| Group | Functions |
| --- | --- |
| File operations | `moveSheet` (rename, relocate, trash the old) · `moveConvertedSheet` (same, keeps the source) · `deleteOldSheet` · `copyFileTemplate` · `moveGetStartedFileToFolder` · `getOrCreateGetStartedFolder` |
| Template + ID resolution | `getTemplateAndsheetIds` · `getTemplateInfo` · `getTemplateIdForSingleSheet` · `fetchIdsMasterData` · `processTemplateAccess` · `findSheetIdAndType` |
| Access | `checkSheetAccess` · `checkTemplateAccess` · `checkFileTemplateAccess` · `checkNewSheetReference` · `getOAuthToken` · `getScopeAuthorizationUrl` · `checkScopePermissions` |
| Versions | `compareSheetVersions` · `checkExportCompatibility` |
| Migration | `prepareImportData` |
| ID writing | `updateIdsMaster` · `getIdsMasterGid` · `updateSheetID` · `updateGetStartedSheetIdsAndReferences` |
| Save-file support | `getSaveFileImportTargets` · `getSaveFileSheetType` |

`getTemplateAndsheetIds` holds the candidate sheet-type list. Ordering matters: a type whose name
is a substring of another must come after it.

---

## 02_SavedFile.js — the save-file parser

`playerInfo.dat` is a GZIP-compressed .NET `BinaryFormatter` stream (NRBF). There is no native
support in Apps Script, so this file implements a reader.

| Symbol | Role |
| --- | --- |
| `labHeaders` … `MasterHeaders` | Eleven header maps, one per category, translating our field names to the game's save-file keys. |
| `parseSaveFileBytes(byteArray)` | Client-callable entry point: ungzip → `parseNRBF` → `extractDataByHeaders` per category → each module's `parse*Data`. Returns `{ parsed, order, failedCategories }`. |
| `parseNRBF(bytes)` | The record loop. Two-pass: collect objects by ID, then `resolve()` follows `_ref` pointers. |
| `unwrapCollection(obj)` | Collapses .NET generics — `List<T>` → array, `Dictionary<K,V>` → object, enum → its numeric `value__`. |
| `utf8Decode` · `blobToUint8Array_` · `bigIntJsonReplacer_` | `Blob.getBytes()` returns *signed* bytes, hence the `& 0xff` mask. `Int64`/`UInt64` return BigInt because coin counters exceed `Number.MAX_SAFE_INTEGER`. |
| `RecordType` · `BinaryType` · `PrimitiveType` · `BinaryArrayType` | NRBF enum tables. |
| `getSaveFilePlayerWaveCapPreference` / `setSaveFilePlayerWaveCapPreference` | The wave-cap choice, persisted per user in `UserProperties`. |

A category that fails to parse is collected into `failedCategories` rather than aborting the whole
parse. Every `parse*Data` must guard with `hasOwnProperty` / null checks so an older save file
missing a field still parses cleanly.

---

## Sheet modules — 03 to 17

| Sheet type | Object | File |
| --- | --- | --- |
| Laboratory | `lab` | `03_Laboratory.js` |
| Workshop | `workshop` | `04_Workshop.js` |
| Ultimate Weapon | `ultimate` | `05_Ultimate_Weapons.js` |
| Themes, Songs & Relics | `themesAndRelics` | `06_Themes_Songs_Relics.js` |
| Bots | `bots` | `07_Bots.js` |
| Vault | `vault` | `09_Vault.js` |
| Cards | `cards` | `10_Cards.js` |
| Modules | `modules` | `11_Modules.js` |
| Guardians | `guardians` | `12_Guardians.js` |
| Player & Stuff | `playerStuff` | `13_Player_&_Stuff.js` |
| IDS Collection | `collection` | `14_IDS_Collection.js` |
| IDS Master | `master` | `15_IDS_Master.js` |
| Effective Paths | `ePaths` | `16_ePaths.js` |
| Relics *(legacy)* | `relics` | `17_Relics.js` |
| Themes & Songs *(legacy)* | `themes` | `17_Themes_&_Songs.js` |

There is no `08_`. The two legacy modules are still readable so old sheets can migrate; nothing
copies those templates any more.

### The contract

Every module implements the same five things, in this order inside the file:

| Member | Role |
| --- | --- |
| `exportData(versionDifference, oldSheetID)` | Dispatch only — look the converter up and call it. |
| `importData(data, newSheetID)` | Read the **new** sheet's layout, build one batch update, flush once. |
| `update*(…)` | Produce the write entries. **Also called directly by `14_IDS_Collection.js`.** |
| `versionN_M(oldSheetID)` | One per template generation; fetches what it needs and delegates. |
| `getVersionN_M*(values)` | Pure readers — no API calls, raw values in, plain objects out. |
| `parse*Data(values)` | Save-file entry point; returns the **same neutral shape** as the converters. |
| `get convertVersionFunctions()` | Getter returning `{ "vN.M": this.versionN_M.bind(this) }`. |
| `isCompatibleVersion(oldVersion)` | Sorts the thresholds descending, returns the newest at or below `oldVersion`, or `null`. |

Export and save-file parsing converge on one neutral object, which is why `importData` serves both
workflows.

Two conventions run through every module:

- **`importData` guards every section it might write.** A partial payload — an older export, or a
  save file lacking a field — updates only what it carries and leaves the rest untouched.
- **Import reads the *new* sheet before writing it.** Row labels are located in the new template's
  own grid, never assumed.

The smallest complete example is `03_Laboratory.js`.

### The two aggregates

**`master` (15_IDS_Master.js)** moves the `IDS` registry — which sheet type lives at which ID — and
the preset names. Its `importData` writes the master's own IDs, which is why a combined update then
only needs the tab's `gid` rather than a full `updateIdsMaster`. From the save file it reads the
game's global preset names, skipping the last (the game stores a dummy there).

**`collection` (14_IDS_Collection.js)** is the single-file arrangement: every category as a tab in
one spreadsheet. Its export produces one object keyed by sheet type; its import delegates to every
other module's `update*` against its own tabs and holds the dropdown named-range tables
(`dvtNamedRangesUW`, `dvtNamedRangesBots`, …) those writes need. Failures are collected per
category into `failedUpdates` rather than aborting. It is by far the largest file, and it carries a
UTF-8 BOM — leave it in place.

### Module-specific notes

- **Laboratory** — the planner is carried as formulas so user planning survives; its tab name
  carries a version suffix, so it is found by substring.
- **Workshop** — everything is per preset, so save-file parsing runs through
  `shared.resolvePresetOrder` before reaching the sheet's fixed slots.
- **Ultimate Weapon** — level cells are dropdowns, so writes go through `shared.getDVTValue`.
- **Modules** — substat effects decode into label/rarity pairs; the inventory is deduplicated,
  keeping the highest-rarity copy of each name and category. `findModuleTypesRowIndex` locates the
  Cannon/Armor/Generator/Core blocks. One converter is written but deliberately not wired into
  `convertVersionFunctions`, waiting on its template.
- **Guardians** — chip levels are dropdowns; writes go through `shared.getDVTValue`.
- **Player & Stuff** — the only category with a client-side transform before import (the wave cap).
- **Vault** — sheets below the current template generation early-return with no data to transfer;
  the older reading code is unreachable and kept as reference.
- **Effective Paths** — the outlier. Four-segment zero-padded versions; detected by the presence of
  its own `eHP`/`eDamage`/`eEcon` tabs rather than by `Home Page!B2`; reads its version through
  `shared.getEPathsVersion`; works in range-relative coordinates via
  `shared.getColumnOffsetFromRange`; has no save-file input at all.

### Enumerating a module's members

```bash
grep -n "^  [a-zA-Z_0-9]*:\|^  get " src/11_Modules.js   # every member of one module
grep -hn "  parse[A-Za-z]*Data: function" src/*.js        # every save-file entry point
grep -ho "^  update[A-Za-z0-9]*:" src/*.js | sort -u      # every import writer
grep -o '"v[0-9][^"]*"' src/11_Modules.js | sort -u       # one module's converter keys
```

---

## The failure envelope

```javascript
{
  success: false,
  code: "SHEET_STRUCTURE",
  expected: false,
  message: "Cards: Could not read required data from spreadsheet",
  reference: "TWR-M4X2K9-A7F3",
  detail: "…",
  trace: ["SheetsAPI.batchGetValues", "importData"],
  outdated: { running, latest, minimum, unsupported },
}
```

`detail` and `trace` belong to the layer that actually failed, not the layer that wrote the
message. `propagate` carries both outward untouched and appends its own frame. `message` is for
the person in the sidebar; `detail` shows only behind *Technical details*.

Envelopes that carry extra fields: `failedUpdates` (on `importData`, per-category failures inside a
multi-category import), `collection: true` (on `getTemplateAndsheetIds` — the sheet has no `IDS`
tab, so it is a Collection not a Master), `versionFiltered: true` (on `getTemplateInfo` — skipped
as already up to date).
