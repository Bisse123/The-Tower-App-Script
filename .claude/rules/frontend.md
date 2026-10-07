---
paths:
  - "src/**/*.html"
---

# Frontend map

No framework. Each page is one HTML file assembled from fragments with `include()`; page state is
module-level variables.

## Pages — client/pages

| Page | Workflow | Bootstrap (`api.js` onload) |
| --- | --- | --- |
| `update.html` | Update Sheet (default route, add-on sidebar) | `gapiLoaded()` |
| `get_started.html` | Get Started | `getStartedGapiLoaded()` |
| `save_file.html` | Import Data From Game; resizes itself as an add-on dialog | `saveFileGapiLoaded()` |
| `consent_dialog.html` | Add-on consent dialog | — |

Shell order: styles in `<head>` → sections → globals `<script>` → fragment scripts → `api.js`.
Globals: `API_KEY`, `APP_ID`, `viewType`, `googleLink` (no trailing slash), the sheet IDs and
`sheetType` (`let` — flows reassign them), and the page flags `getStarted` / `saveFile`.

## client/common

`header_*` (branding, click-to-copy) · `status_*` (status line, `detectMobile`) · `error_*`
(`AppError`, `runAppsScript`; scripts use `includeTemplate` to inline the error contract) ·
`version_banner_scripts` · `consent_*` · `templates_scripts` (`SHEET_TEMPLATES`, the template-ID
registry) · `base_styles`.

## client/update

| Scripts | Holds |
| --- | --- |
| `state` | Every flag and file bucket — loads before `status_scripts`, which sets `isMobile` |
| `picker` · `access` · `copy` | Picker, access checks, template copying and combined-update creation |
| `select_import` | Manual file pickers when parameters are missing |
| `legacy_themes` | Merging pre-v4.0 Themes & Songs + Relics into one sheet |
| `transfer` | Export/import fan-outs, both conversions, `stageMasterExport` |
| `move` · `results` | The move step; the results summary |
| `flows` | Entry buttons, `showContinueSection`, `renderIdMasterOptions`, button states, resets |
| `authorize` | `authorizeAndContinue` and the sidebar context |

Flow flags: `isUpdateSingleSheetFlow` · `isCombinedUpdate` · `isConvertToMasterFlow` ·
`isConvertToCollectionFlow` · `hasSubsheetsToUpdate` · `mergeLegacyThemesSheets` ·
`masterIdsWritten`. File buckets run `copiedTemplateFiles` → `exportedFiles*` → `importedFiles*` →
`movedFiles`; `prepareImportData` takes the pending buckets, so Import retries what did not land.

## client/get_started

`state` · `folder` (setup, folder picker) · `copy` · `id_links` (cross-linking) · `results`
(rendering, retry, `sanitizeGetStartedUrl`) · `authorize` (`authorizeGetStarted`).

## client/save_file

Scripts: `state` · `parse` (file, Drive picker) · `view` (category cards, selection, view toggle,
sheet prefetch) · `access` (targets, outdated checks) · `diff_primitives` (`sf*`) · `diff`
(`render*Diff`) · `body` (`render*Body`) · `wave_cap` · `import` · `authorize` (`authorizeSaveFile`).
Styles `layout` → `body` → `diff` must stay in that order. `guide_*` is the playerInfo.dat guide.

## Cycles every page shares

- **Access:** under `drive.file` a known ID is not enough. Check access → picker seeded with
  `setFileIds` (needs `setOrigin(google.script.host.origin)`) → re-check → continue. Not
  owned/editable is a hard stop; editable but not owned is an amber warning.
- **Consent:** web app opens a popup and polls every 700 ms for it closing; the sidebar opens
  `consent_dialog`, which sets a `UserProperties` signal the sidebar polls every 1 200 ms
  (`consume…` reads and deletes it). Consent then calls the page's authorize function by flag:
  `saveFile` → `authorizeSaveFile`, `getStarted` → `authorizeGetStarted`, else
  `authorizeAndContinue`. `save_file.html` never declares `getStarted`; the `saveFile` branch must
  come first. The dialog declares its own `authorizeAndContinue`.
