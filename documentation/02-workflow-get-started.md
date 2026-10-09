# 02 — Get Started

Copies every template into the user's Drive and links their IDs, leaving a working set of sheets.

| | |
| --- | --- |
| Add-on | `Import Data ▸ Get Started` → `showGetStartedDialog()` |
| Web app | `?page=getstarted` |
| Page | `client/pages/get_started.html`, scripts in `client/get_started/` |
| Server | `server/workflows/get_started.js`, `copyFileTemplate` |

## What the user sees

An explainer, a choice between **IDS Master and subsheets** (Master plus ten subsheets) and **IDS
Collection** (one file), a **Copy Templates** button, and manual copy links. An Effective Paths sheet
is copied either way. Template IDs come from `SHEET_TEMPLATES`.

In the add-on, when the open spreadsheet is already an Effective Paths sheet, that sheet is copied
instead of the Effective Paths template.

## Flow

```mermaid
sequenceDiagram
    participant C as Client
    participant S as Server
    C->>S: getOrCreateGetStartedFolder()
    Note over S: finds or creates "The Tower" (a new one is shared by link)
    C->>S: checkTemplateAccess ×N
    opt any inaccessible
        C->>C: picker → re-check
    end
    par each template
        C->>S: copyFileTemplate(…, folderID)
    end
    par each copy
        C->>S: updateGetStartedSheetIdsAndReferences(fileId, type, relatedIDs)
        Note over S: writes the IDs, renames to "<type> <version>"
    end
    C->>C: render links; offer "Retry Failed Copies"
```

## ID linking

| Mode | File | Receives |
| --- | --- | --- |
| Master + subsheets | IDS Master | every subsheet's ID (`updateIdsMaster`) |
| | each subsheet, Effective Paths | the Master's ID |
| Collection | IDS Collection | its own ID |
| | Effective Paths | the Collection's ID, as its Master |

A failed rename is logged and does not fail the linking.

## Retry

The client keeps what it created and what failed. **Retry Failed Copies** re-runs only the failed
copies and links, recomputing the links from everything created so far — and re-links the Master
whenever a subsheet is retried, so it learns the newcomer's ID.
