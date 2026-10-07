/**
 * Configuration & Google Sheet Triggers for Backlog Sync
 * 
 * IMPORTANT SECURITY NOTICE:
 * This file contains sensitive credentials (API key, domain, email mapping).
 * Keep this file private and do NOT commit it to public version control repositories.
 * 
 * SETUP INSTRUCTIONS:
 * 1. In your Google Sheet, open Extensions > Apps Script.
 * 2. Add your library under Libraries (+) with identifier: BacklogSync
 * 3. Paste this file as `config.js` (or `Code.gs`).
 * 4. Refresh your Google Sheet to see the "Backlog Sync" menu.
 */

const CONFIG = {
  // Your Backlog domain (e.g. "yourcompany.backlog.com" or "yourcompany.backlog.jp")
  // You can include or omit "https://"
  SPACE_DOMAIN: "YOUR_SPACE.backlog.com",

  // API Key generated in Backlog: Personal Settings > API > Register New API Key
  API_KEY: "YOUR_BACKLOG_API_KEY",

  // Your Backlog Project Key (e.g. "IZ" or "PROJ")
  PROJECT_KEY: "YOUR_PROJECT_KEY",

  // Default Issue Type Name (e.g. "Task" or null to use the first issue type in the project)
  DEFAULT_ISSUE_TYPE: "Task",

  // Priority ID: 2 = High, 3 = Normal, 4 = Low
  PRIORITY_ID: 3,

  // Row index where table headers are located (data starts below this row)
  HEADER_ROW_INDEX: 5,

  // Column numbers (1-indexed):
  // Col A = 1 (empty), Col B = 2, Col C = 3, Col D = 4, Col E = 5, Col F = 6
  COLUMNS: {
    TASK: 2,         // Column B (Task Summary / Issue Title)
    ASSIGNED_TO: 3,  // Column C (Assignee Name)
    START_DATE: 5,   // Column E (Start Date)
    END_DATE: 6      // Column F (End Date / Due Date)
  }
};

/**
 * Mapping of Assignee Name (as shown in the sheet) to their Backlog Email Address.
 * Add all team members here.
 */
const USER_EMAIL_MAP = {
  // "John Doe": "john.doe@yourcompany.com",
  // "Jane Smith": "jane.smith@yourcompany.com",
};

// ============================================================================
// SHEET UI TRIGGERS & HANDLERS
// ============================================================================

/**
 * Automatically creates the custom menu when opening the Google Sheet.
 */
function onOpen() {
  BacklogSync.createMenu();
}

/**
 * Triggered when clicking "Backlog Sync > Sync Selected Row" in Google Sheets.
 * Reads the selected row and calls the library to create the issue in Backlog.
 */
function syncSelectedRowToBacklog() {
  BacklogSync.syncSelectedRowToBacklog(CONFIG, USER_EMAIL_MAP);
}

// Node / CommonJS module export for testing or bundling (optional)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { CONFIG, USER_EMAIL_MAP, onOpen, syncSelectedRowToBacklog };
}
