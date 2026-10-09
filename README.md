# The Tower — App Script

Google Apps Script project behind the community spreadsheets ("IDS Sheets™") for the mobile game
**The Tower**. One script project serves two faces with the same server functions:

- a **Google Sheets add-on** — the `Import Data` menu inside a sheet;
- a **web app** — `doGet`, run as the visiting user.

## What it does

| Workflow | Does | Docs |
| --- | --- | --- |
| **Get Started** | Copies every template into a `The Tower` Drive folder and cross-links their IDs | [02](documentation/02-workflow-get-started.md) |
| **Update Sheet** | Copies a new template version, moves the user's data into it, then swaps it in for the old sheet | [03](documentation/03-workflow-update-sheets.md) |
| **Import Data From Game** | Reads the game's `playerInfo.dat` save file and writes it into the sheets | [04](documentation/04-workflow-save-file-import.md) |

## The sheets

Players keep their data in one of two interchangeable arrangements, and either feeds an
**Effective Paths** calculation sheet:

- **IDS Master + subsheets** — one file per category (Laboratory, Workshop, Ultimate Weapon,
  Themes Songs & Relics, Bots, Vault, Cards, Modules, Guardians, Player & Stuff); the Master's
  `IDS` tab records each subsheet's ID, template link and version.
- **IDS Collection** — every category as a tab in one file.

The app converts between the two in either direction.

## Layout

```
src/                    everything clasp pushes to Apps Script
  appsscript.json       manifest: scopes, advanced services, web-app settings
  server/               core · helpers · entry · workflows · savefile · sheets/<type>
  client/               pages · common · update · get_started · save_file
documentation/          one document per area
docs/                   save-format reference JSON (git-ignored)
.github/                deploy workflow and its scripts
bump.js · check-src.js  version bumper · local checks
```

The client orchestrates; the server is stateless single-purpose functions reached through
`google.script.run`, which return failures rather than throwing. Every sheet type exports and
imports through one neutral data shape, so migrating a sheet and importing a save file share the
same importer. See [01 — Architecture](documentation/01-architecture.md).

## Authorization

The app requests `drive.file`: it can only open files the user hands it through the Google Picker,
so every workflow checks access, opens the picker, re-checks and continues. The Picker needs the
`API_KEY` and `APP_ID` script properties.

## Development

```bash
npm install
clasp login            # an account with access to both script projects
npm run check          # local checks
npm run sandbox        # push to the dev project
npm run dev            # push to production HEAD — where a change is tested before merging
```

Merging to `main` deploys, but only if production HEAD equals the merged code. Details:
[07 — Deployment](documentation/07-deployment.md). All docs: [documentation/](documentation/README.md).
