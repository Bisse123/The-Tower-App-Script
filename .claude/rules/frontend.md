---
paths:
  - "src/*.html"
  - "src/**/*.html"
---

# Frontend map

No framework, no store, no build. Apps Script serves one HTML file per page, assembled from
numbered fragments by `include()`. All page state is module-level variables at the top of the
fragment scripts.

---

## Pages

| Page shell | Workflow | Bootstrap |
| --- | --- | --- |
| `20_WebApp.html` | Update Sheet (sidebar + web app default route) | `gapiLoaded()` |
| `20_getStartedApp.html` | Get Started (modal + `?page=getstarted`) | `getStartedGapiLoaded()` |
| `20_SavedFileApp.html` | Import Data From Game (modal + `?page=savefile`) | `saveFileGapiLoaded()` |
| `29_addon_consent_dialog.html` | Sidebar-only consent dialog (whole page, no fragments) | — |

Each shell ends with `<script src="https://apis.google.com/js/api.js" onload="…">` and differs only
in its bootstrap function.

Fragment order inside a shell is load-bearing: **styles in `<head>` → sections → the
server-injected globals `<script>` → fragment scripts.** Fragment scripts run at parse time and
read those globals immediately.

`20_SavedFileApp.html` also resizes itself when opened as an add-on dialog, via
`google.script.host.setWidth/setHeight` clamped to 80–85 % of the screen.

---

## Fragments

Fragments come in triples: `_section` (markup), `_styles` (a `<style>` block), `_scripts` (a
`<script>` block).

| Fragment | S | Y | J | Purpose | Used by |
| --- | :-: | :-: | :-: | --- | --- |
| `21_header` | ✓ | ✓ | ✓ | Branding, creator-code chips, shared click-to-copy | all |
| `21_shared` | | ✓ | ✓ | Picker, access checks, template copying, combined update | styles: all · **scripts: WebApp only** |
| `21_templates` | | | ✓ | `SHEET_TEMPLATES` — the template-ID registry | all |
| `22_status` | ✓ | ✓ | ✓ | The one-line status bar; mobile detection | all |
| `22_error` | ✓ | ✓ | ✓ | The error panel, `AppError`, `runAppsScript` | all |
| `23_getStarted` | ✓ | ✓ | ✓ | Get Started explainer and quick setup | getStartedApp |
| `24_selectImport` | ✓ | ✓ | ✓ | Manual file selection when params are missing | WebApp |
| `25_fileAccess` | ✓ | ✓ | ✓ | The update workflow's buttons and orchestration | WebApp |
| `26_instructions` | ✓ | ✓ | | Picker usage tips, desktop vs. mobile | WebApp |
| `27_consent` | ✓ | ✓ | ✓ | Missing-scope modal and recheck loop | all + dialog |
| `28_saveGuide` | ✓ | ✓ | | Where to find `playerInfo.dat` — pure markup | SavedFileApp |
| `28_saveFile` | ✓ | ✓ | ✓ | Parse, diff, select, import | SavedFileApp |

`21_shared_scripts.html` is included **only** by `20_WebApp.html`. The other two pages carry their
own narrower equivalents. `22_error_scripts` is the only fragment included with `includeTemplate`,
because its scriptlet inlines the error contract from `ERROR_DEFS`.

---

## Server-injected globals

Set by Apps Script templating in the shell:

```javascript
const API_KEY   = "<?= API_KEY ?>";    // ScriptProperties — Picker developer key
const APP_ID    = "<?= APP_ID ?>";     // ScriptProperties — Picker app ID
const viewType  = "<?= viewType ?>";   // "webapp" | "sidebar" | "addon-consent-dialog"
const googleLink = "https://docs.google.com/spreadsheets/d";
let newSheetID, oldSheetID, idMasterID, sheetType;   // let — the flows reassign them
const getStarted = false;   // page flags read by the shared consent script
const saveFile   = false;
```

The sheet IDs are `let` on purpose: flows reassign them freely
(`newSheetID = copyResult.fileId`, `sheetType = "IDS Master"` mid-conversion).

`googleLink` has **no trailing slash** and lives in the shell, not a fragment. A literal `//`
inside an included fragment's `<script>` is stripped as a comment and kills the whole block, so
URLs must be assembled from this constant.

> `20_SavedFileApp.html` declares `saveFile = true` but **does not declare `getStarted`**. The
> `saveFile` branch in `27_consent_scripts` short-circuits before `getStarted` is read. Reordering
> those branches turns it into a `ReferenceError`.

---

## Calling the server

`22_error_scripts` defines `runAppsScript`, the one wrapper every page uses. It rejects with a
**normalised** error — the same shape a failed envelope has. Bulk operations are plain
`Promise.all` fan-outs. Which call to make, and the fan-out rules:
[conventions.md ▸ Errors — client](conventions.md#errors--client).

### `AppError`

`show` · `showAll` · `surfaceBatch` · `check` · `log` · `clear` · `normalize` · `recognise` ·
`isForeign` · `finalize`. Tables: `CODES` · `EXPECTED` · `TITLES` · `RECOGNISED`.
`ERROR_CONTRACT.MESSAGES` carries the inlined server messages.

`normalize` is idempotent — normalising an already-normalised error extends its trace instead of
reclassifying it. `window.onerror` and `unhandledrejection` are wired in; cross-origin
`"Script error."` is ignored, and `isForeign` drops anything raised by a browser extension.

Two codes the page classifies itself from the failure text: `AUTH_UNAVAILABLE` (no Google session,
blocked cookies, refused popup) and `NETWORK_BLOCKED` (request never reached Google).

Also here: `versionNotice(status)` and `checkAppVersion()`, which drive the outdated banner.

---

## Where each workflow's logic lives

### Update Sheet — `21_shared_scripts` + `25_fileAccess_scripts` + `24_selectImport_scripts`

`21_shared_scripts`: `gapiLoaded` · `isSheetId` · `usableSheetId` · `checkValidParameters` ·
`hideAllSections` · `calculatePickerSize` · `openMultiFilePicker` · `openSingleFilePicker` ·
`formatSheetTypeList` · `describeBlockedTemplates` · `blockedTemplatesNoticeHtml` ·
`describeNothingToCopy` · `onTemplateAndOldSheetAccessCheckFailure` · `onSheetAccessCheckSuccess` ·
`onSheetAccessCheckFailure` · the `UpdateSingleSheetButton` enable/disable/show/hide/reset set ·
`proceedWithUpdateSingleSheetTemplateCopy` · `proceedWithTemplateCopying` ·
`proceedWithCombinedUpdate`.

`25_fileAccess_scripts`: the legacy-themes trio (`detectLegacyThemesSheets` ·
`mergeLegacyThemesTemplates` · `expandMergedTemplateFiles` · `collapseMergedImportTasks`) ·
`showFileAccessSection` · `recordPreparationFailures` · `remapMasterIdsToNewSheets` ·
`onprepareImportDataFailure` · `onImportFailure` · `moveSheet` · `surfaceFailureIfBug` ·
`escapeSummaryHtml` · `summaryFileRow` · `displayOperationResults` · `movedSheetLink` ·
`onmoveSheetFailure` · `showAccessibleFilesSummary` · `showNotOwnedWarning` ·
`showContinueSection` · `updateSubsheetsOnly` · the `IDSMasterButtons` set ·
`resetIDSMasterState` · `resetToCollectionState`.

`24_selectImport_scripts`: `showSelectImportSection` and the manual-picker chain —
`checkFileTemplateAccessSuccess` / `Failure` · `copyFileTemplateSuccess` / `Failure` ·
`showFileCopiedSuccess` · `checkNewSheetRef` · `findAndGrantNewSheet` · `findAndGrantIdMaster` ·
`findAndGrantOldSheet`.

**The five flows** are selected by flags: `isUpdateSingleSheetFlow` · `isCombinedUpdate` ·
`isConvertToMasterFlow` · `isConvertToCollectionFlow` · `hasSubsheetsToUpdate` ·
`mergeLegacyThemesSheets` · `masterIdsWritten`.

**The state buckets** are a left-to-right pipeline — `copiedTemplateFiles` →
`exportedFilesSuccess` / `exportedFilesFailed` → `importedFilesSuccess` / `importedFilesFailed` →
`movedFiles`. `prepareImportData` takes all three pending buckets, so pressing Import again retries
exactly what did not land. `showContinueSection` and `renderIdMasterOptions` are the two
visibility switchboards.

### Get Started — `23_getStarted_scripts`

`renderGetStartedTemplateLinks` · `continueGetStarted` · `updateGetStartedSelectedFolderDisplay` ·
`getStartedGapiLoaded` · `renderGetStartedCopyResult` · `copyGetStartedTemplateAsync` ·
`mergeKnownCreatedFile` · `processGetStartedIDUpdateResults` · `processGetStartedCopyResults` ·
`applyGetStartedIDUpdates` · `retryFailedCopies`.

State is the chosen Drive folder, the copy mode, and the created/failed lists that drive the retry
model. `retryFailedCopies` re-derives relationships from `allCreatedFiles`, so a subsheet copied on
attempt 2 still links to a master copied on attempt 1.

### Import Data From Game — `28_saveFile_scripts`

Parse and setup: `parseFile` · `parseSaveFileByteArray` · `openSaveFileDrivePicker` ·
`clearSaveFileParseUI` · `updateHeldSaveFileLabel` · `listSaveFileNames` ·
`renderSaveFileParseFailures` · `renderSaveFileResults` · `setSaveFileSetupError`.

Targets and gating: `sfTypeChecked` · `sfTypeOutdatedInfo` · `getSaveFileOutdatedTargets` ·
`saveFileTargetLabel` · `refreshSaveFileImportSummary` · `captureSaveFileSelection` ·
`updateSaveFileViewControls` · `buildSaveFileCategoryBody` · `startSaveFileSheetPrefetch` ·
`runSaveFileImport` · `sfImportPayload` · `sfPlayerWaveData`.

Diff primitives: `sfNormLevel` · `sfNormGeneric` · `sfNormBool` · `sfTrimLevel` · `sfDisplayLevel` ·
`sfDisplayValue` · `sfDisplayBool` · `sfDiffRow` · `sfDiffAddItem` · `sfDiffRemItem` · `sfDiffGrid` ·
`sfDiffGridStacked` · `sfDiffBlock` · `sfDiffSub` · `sfDiffNoneHtml`.

Per-category renderers, dispatched by `renderSaveFileDiffBody`: `renderLaboratoryDiff` ·
`renderWorkshopDiff` · `renderUltimateDiff` · `renderThemesAndRelicsDiff` · `renderThemesDiff` ·
`renderRelicsDiff` · `renderBotsDiff` · `renderVaultDiff` · `renderCardsDiff` · `renderModulesDiff` ·
`renderModulesPresetsDiff` · `renderModulesInventoryDiff` · `renderGuardiansDiff` ·
`renderPlayerDiff`.

`sfDiffNoneHtml()` emits a marker the card renderer detects to mark a category "✓ no differences"
and untick it. A per-category export failure lands as `{ __error: "…" }` and renders as a notice
rather than failing the page.

---

## The two cycles every page inherits

### The access-grant cycle

Under `drive.file`, knowing a sheet ID is not enough — the user must select the file in the Google
Picker, which grants access to that specific file. So: **check access → open picker → re-check →
continue**, everywhere.

Buckets: `accessibleFilesCache` · `inaccessibleFilesCache` (needs a grant) ·
`notOwnedFilesCache` (**blocking** — user is neither owner nor editor) · `warningFilesCache`
(editable but not owned; amber, non-blocking) · `templateAccessibleFilesCache` /
`templateInaccessibleFilesCache`.

Picker specifics: the view is seeded with `setFileIds(...)` so only files needing a grant are
offered; `setOrigin(google.script.host.origin)` is required to render inside the Apps Script
iframe; on mobile `google.picker.Feature.NAV_HIDDEN` is enabled. `pickerCallback` is a large
branch on the current flow, because a picker completion has to resume whichever step opened it.

### The consent flow

Scopes are not granted by opening the page, and a sidebar cannot host the cross-origin consent
screen — so the two views differ:

- **webapp** → `showConsentModal(url)` → `window.open` → poll every 700 ms for the popup closing →
  `recheckAccessToken(true)`.
- **sidebar** → `showAddonConsentDialog(url)` opens `29_addon_consent_dialog` → the user authorizes
  → `markAddonConsentReadySignal()` → the sidebar polls `consumeAddonConsentReadySignal()` every
  1 200 ms → the dialog closes itself.

The signal is a timestamp in `UserProperties`; `consume…` reads **and deletes** it, so it fires
exactly once.

`27_consent_scripts.html` is shared by all three pages *and* by the dialog, which overrides
`authorizeAndContinue` and stubs `setStatusText` / `setStatusWithSpinner` onto its own badge.
`consentScopeCheckOutcome` is `"idle" | "granted" | "missing_scopes" | "error"`; `missing_scopes`
keeps the manual button hidden, `error` reveals *Recheck scope access*.

---

## Status, mobile and escaping

One status line drives every page: `setStatusWithSpinner(msg)` while working,
`setStatusText(msg)` for a terminal state. `AppError._render` drives it too — with the panel on
the page the status line takes the error's title and the panel carries the message; on a page
with no panel the status line takes the message instead.

`detectMobile()` combines user-agent keywords, `ontouchstart`/`maxTouchPoints` and a ≤768 px width
check, then swaps the instruction panel and substitutes shorter status strings.

Names and URLs reaching `innerHTML` must be escaped — see
[conventions.md ▸ Escaping](conventions.md#escaping).

---

## Two structural quirks

Buttons live in `_section.html` with `style="display: none"` and `disabled`, and are revealed by
the flow. `copyButton`, `copyUpdateButton` and `updateMasterOnlyButton` are superseded flows kept
commented-out for reference; the live element is the one without comment markers.

`21_header_scripts.html` holds the shared clipboard helper: `navigator.clipboard`, falling back to
`document.execCommand("copy")` (the Clipboard API is blocked in some sandboxed Apps Script
iframes), and finally selecting the text so `Ctrl+C` still works.
