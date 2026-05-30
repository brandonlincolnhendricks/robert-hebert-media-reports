# RHM Weekly Report Generator — Apps Script Source

This folder is the clasp-managed source of truth for the bound Apps Script project
that powers the RHM Weekly Report Generator.

## One-time setup (already done if `.clasp.json` exists in this folder)

1. Enable the Apps Script API for your Google account:
   https://script.google.com/home/usersettings → toggle ON.
2. Authenticate clasp locally:
   ```bash
   clasp login
   ```
3. Get the script ID. Open the bound spreadsheet → Extensions → Apps Script.
   The URL looks like `https://script.google.com/d/SCRIPT_ID_HERE/edit` — copy
   the `SCRIPT_ID_HERE` portion.
4. From this folder, write `.clasp.json`:
   ```json
   { "scriptId": "<paste here>", "rootDir": "." }
   ```

## Deploy

From this folder:

```bash
clasp push --force
```

That overwrites the remote project with `Code.js` + `appsscript.json` from this
directory. `--force` is needed any time the manifest changes.

## Pull (if someone edited in the web UI)

```bash
clasp pull
```

Pulls the live remote into this folder. Review the diff before committing.

## Open the live project

```bash
clasp open
```

## What gets pushed

`.claspignore` restricts the push to exactly two files:

- `Code.js` (renamed to `Code.gs` server-side automatically)
- `appsscript.json`

Nothing else in the repo is touched.
