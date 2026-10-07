# Backlog Google Sheets Sync

This Google Apps Script syncs task rows from a Google Sheet directly into **Backlog** issues with assignee resolution, date formatting, and custom menu integration.

---

## Project Structure

```
.
├── sync.js            # THE LIBRARY: Core sync logic & Backlog API integration (deploy this as the library)
├── config.js          # CONSUMER CODE: Configuration, assignee mapping, and onOpen() menu trigger
├── config.example.js  # Safe template to share / commit
├── .gitignore         # Ignores config.js to prevent leaking credentials
└── README.md          # Setup and library documentation
```

---

## 1. Standalone Usage (Single Apps Script Project)

If both `config.js` and `sync.js` live inside your Google Spreadsheet's Apps Script editor:

1. Open your Google Sheet and navigate to **Extensions > Apps Script**.
2. Add a new script file named `config.js` (or `Config.gs`) and paste the contents of `config.js`. Fill in your `SPACE_DOMAIN`, `API_KEY`, `PROJECT_KEY`, and `USER_EMAIL_MAP`.
3. Add `sync.js` (or `Sync.gs`) with the contents of `sync.js`.
4. Refresh the Google Sheet. The custom menu **Backlog Sync > Sync Selected Row** will appear.
5. Select any task row in your sheet and click **Backlog Sync > Sync Selected Row**.

---

## 2. Using as a Google Apps Script Library

You can deploy `sync.js` as an Apps Script Library so multiple spreadsheets can reuse the exact same codebase without duplicating code.

### Step 1: Deploy the Library Project
1. Create a standalone Google Apps Script project (e.g. named `BacklogSyncLibrary`) containing `sync.js`.
2. Click **Deploy > New deployment**.
3. Select type **Library**, give it a description (e.g. `v1.0.0`), and deploy.
4. Copy the **Script ID** from **Project Settings**.

### Step 2: Include the Library in Your Sheet's Script
1. Open your target Google Sheet's Apps Script editor (**Extensions > Apps Script**).
2. Beside **Libraries**, click **+** (Add a library).
3. Paste the Script ID, select the latest version, and set the identifier as **`BacklogSync`**.
4. In your sheet's script, paste the entire contents of `config.js`:

```javascript
// config.js (inside consumer Google Sheet)
const CONFIG = {
  SPACE_DOMAIN: "yourcompany.backlog.com",
  API_KEY: "YOUR_API_KEY",
  PROJECT_KEY: "PROJ",
  DEFAULT_ISSUE_TYPE: "Task",
  PRIORITY_ID: 3,
  HEADER_ROW_INDEX: 5,
  COLUMNS: {
    TASK: 2,
    ASSIGNED_TO: 3,
    START_DATE: 5,
    END_DATE: 6
  }
};

const USER_EMAIL_MAP = {
  "Bismillakhan": "bismillakhan@yourcompany.com",
};

// UI Triggers calling the library
function onOpen() {
  BacklogSync.createMenu();
}

function syncSelectedRowToBacklog() {
  BacklogSync.syncSelectedRowToBacklog(CONFIG, USER_EMAIL_MAP);
}
```

---

## 3. Alternative: Secure Secrets via Script Properties

If you want to avoid storing sensitive API keys in code files altogether:

1. In the Apps Script editor, go to **Project Settings > Script Properties**.
2. Add the following properties:
   - `BACKLOG_SPACE_DOMAIN` : e.g. `yourcompany.backlog.com`
   - `BACKLOG_API_KEY` : your Backlog API Key
   - `BACKLOG_PROJECT_KEY` : e.g. `PROJ`
   - `BACKLOG_USER_EMAIL_MAP` : JSON string (e.g. `{"Bismillakhan":"bismillakhan@yourcompany.com"}`)
3. Then call:

```javascript
function syncSelectedRowToBacklog() {
  const { config, userEmailMap } = BacklogSync.loadConfigFromScriptProperties();
  BacklogSync.syncSelectedRowToBacklog(config, userEmailMap);
}
```
