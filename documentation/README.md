# Documentation index

Deep-dive documentation for **The Tower — App Script**. Start with the
[root README](../README.md) for the overview; these documents assume it.

| # | Document | Read it when you… |
| --- | --- | --- |
| 01 | [Architecture](01-architecture.md) | …need to understand caching, the Sheets/Drive wrappers, or how the app finds anything inside a spreadsheet. |
| 02 | [Get Started workflow](02-workflow-get-started.md) | …are changing template copying, the `The Tower` folder, or ID cross-linking. |
| 03 | [Update Sheets workflow](03-workflow-update-sheets.md) | …are touching the migration flow — the biggest and most stateful part of the app. |
| 04 | [Save-file import workflow](04-workflow-save-file-import.md) | …are adding a game field, fixing the `.dat` parser, or changing the diff view. |
| 05 | [Sheet modules reference](05-sheet-modules.md) | …need to add a new template version or a new sheet type. |
| 06 | [Frontend](06-frontend.md) | …are editing the HTML pages, the picker, the consent flow, or client state. |
| 07 | [Deployment & operations](07-deployment.md) | …are shipping a release or debugging CI. |
| 08 | [Error handling & reporting](08-error-handling.md) | …are diagnosing a reported failure, or writing code that can fail. |

## Common tasks → where to look

| Task | Document | Section |
| --- | --- | --- |
| A new template version of an existing sheet was released | [05](05-sheet-modules.md) | *Adding a new version converter* |
| A brand-new sheet type must be supported | [05](05-sheet-modules.md) | *Adding a new sheet type* |
| The game added a new stat to the save file | [04](04-workflow-save-file-import.md) | *Adding a new save-file field* |
| "Sheet not found" / "Could not find sheet ID for X" | [01](01-architecture.md) | *Discovery by label scanning* |
| Stale data after an import | [01](01-architecture.md) | *Cache invalidation* |
| A picker keeps re-asking for access | [06](06-frontend.md) | *The access-grant cycle* |
| The deploy workflow fails with "Production HEAD does not match" | [07](07-deployment.md) | *The CI guard* |
| A user quotes an error reference like `TWR-M4X2K9-A7F3` | [08](08-error-handling.md) | *Runbook* |
| Cloud Logging / Error Reporting shows nothing | [08](08-error-handling.md) | *Google Cloud* |
