/**
 * Google Apps Script: Sync Selected Row to Backlog
 * 
 * Features:
 * - Adds a custom menu item in Google Sheets ("Backlog Sync > Sync Selected Row")
 * - Reads task title, assignee, start date, and end date from the currently selected row
 * - Column Mapping:
 *     Column A: Empty (ignored)
 *     Column B: Task Summary / Issue title
 *     Column C: Assigned To (Name)
 *     Column D: Ignored
 *     Column E: Start Date
 *     Column F: End Date / Due Date
 * - Resolves the assignee name to email and gets Backlog User ID automatically
 * - Normal Priority (ID: 3)
 * - Backlog REST API v2 using application/x-www-form-urlencoded
 * - Does NOT write anything back to the sheet; notifies via Toast & Alert popup
 */

// ============================================================================
// CONFIGURATION
// ============================================================================
const CONFIG = {
  // Your Backlog domain (e.g. "yourcompany.backlog.com" or "yourcompany.backlog.jp")
  // You can include or omit "https://"
  SPACE_DOMAIN: "team-o.backlog.com",

  // API Key generated in Backlog: Personal Settings > API > Register New API Key
  API_KEY: "cBZwHKPSKQE0aHTXPorHez9uITRMnhyqSpw1R5PncJatd6XbONSqj5IZol5xtRUX",

  // Your Backlog Project Key (e.g. "IZ" or "PROJ")
  PROJECT_KEY: "TEAMO",

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
  "Ashwin Sanalkumar": "imask0110@gmail.com",
  // "John Doe": "john.doe@yourcompany.com",
  // "Jane Smith": "jane.smith@yourcompany.com",
};

// ============================================================================
// MENU TRIGGER
// ============================================================================

/**
 * Automatically creates the custom menu when opening the Google Sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Backlog Sync')
    .addItem('Sync Selected Row', 'syncSelectedRowToBacklog')
    .addToUi();
}

// ============================================================================
// MAIN SYNC FUNCTION
// ============================================================================

/**
 * Reads the currently selected row and creates an issue in Backlog.
 */
function syncSelectedRowToBacklog() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const rowIndex = activeCell.getRow();

  // 1. Validation: Prevent running on header or title rows
  if (rowIndex <= CONFIG.HEADER_ROW_INDEX) {
    ui.alert(
      'Invalid Selection',
      `Please select a valid task row below row ${CONFIG.HEADER_ROW_INDEX}.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // 2. Read row data from Columns B to F
  const maxCol = Math.max(...Object.values(CONFIG.COLUMNS));
  const rowValues = sheet.getRange(rowIndex, 1, 1, maxCol).getValues()[0];

  const summary = String(rowValues[CONFIG.COLUMNS.TASK - 1] || '').trim(); // Column B
  const assigneeName = String(rowValues[CONFIG.COLUMNS.ASSIGNED_TO - 1] || '').trim(); // Column C
  const startDateRaw = rowValues[CONFIG.COLUMNS.START_DATE - 1]; // Column E
  const endDateRaw = rowValues[CONFIG.COLUMNS.END_DATE - 1]; // Column F

  // 3. Validation: Check if row is empty or a section header (e.g. "Backend Development")
  if (!summary) {
    ui.alert('Empty Task', 'The selected row does not have a task title in Column B.', ui.ButtonSet.OK);
    return;
  }

  if (!assigneeName && !startDateRaw && !endDateRaw) {
    ui.alert(
      'Section Header Detected',
      `"${summary}" appears to be a section header rather than an individual task. Please select a row with a task, assignee, or dates.`,
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Connecting to Backlog...', 'Please wait', 5);

    // 4. Get Project Info & Issue Type ID from Backlog
    const project = getBacklogProject(CONFIG.PROJECT_KEY);
    if (!project || !project.id) {
      throw new Error(`Could not find project or numeric ID for "${CONFIG.PROJECT_KEY}".`);
    }

    const issueTypeId = getBacklogIssueTypeId(project.id, CONFIG.DEFAULT_ISSUE_TYPE);

    // 5. Resolve Assignee ID via email mapping
    let assigneeId = null;
    if (assigneeName) {
      assigneeId = resolveAssigneeId(project.id, assigneeName);
      if (!assigneeId) {
        const proceed = ui.alert(
          'Assignee Not Found',
          `Could not match "${assigneeName}" to any project member in Backlog.\n\nWould you like to create the task without an assignee?`,
          ui.ButtonSet.YES_NO
        );
        if (proceed !== ui.Button.YES) {
          return;
        }
      }
    }

    // 6. Format Dates (YYYY-MM-DD)
    const timeZone = Session.getScriptTimeZone();
    const startDate = formatDate(startDateRaw, timeZone);
    const dueDate = formatDate(endDateRaw, timeZone);

    // 7. Build Backlog Issue Payload
    const payload = {
      projectId: project.id,
      summary: summary,
      issueTypeId: issueTypeId,
      priorityId: CONFIG.PRIORITY_ID
    };

    if (assigneeId) payload.assigneeId = assigneeId;
    if (startDate) payload.startDate = startDate;
    if (dueDate) payload.dueDate = dueDate;

    Logger.log('Payload to submit: ' + JSON.stringify(payload));

    // 8. Send POST Request to Backlog
    const createdIssue = createBacklogIssue(payload);
    const issueUrl = `${getBaseUrl()}/view/${createdIssue.issueKey}`;

    // 9. Feedback to user (does not touch or modify the sheet)
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Issue ${createdIssue.issueKey} created successfully!`,
      'Success',
      5
    );

    ui.alert(
      'Issue Created Successfully',
      `Backlog Issue: ${createdIssue.issueKey}\nSummary: ${summary}\n\nURL: ${issueUrl}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log('Error: ' + error.stack);
    ui.alert('Error Syncing to Backlog', error.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// BACKLOG API HELPERS
// ============================================================================

/**
 * Returns clean base URL without trailing slash or protocol duplicates.
 */
function getBaseUrl() {
  const cleanDomain = String(CONFIG.SPACE_DOMAIN || '')
    .trim()
    .replace(/^https?:\/\//i, '')
    .replace(/\/+$/, '');
  return `https://${cleanDomain}`;
}

/**
 * Fetches project details from Backlog.
 */
function getBacklogProject(projectKey) {
  const cleanKey = String(projectKey || '').trim();
  const url = `${getBaseUrl()}/api/v2/projects/${encodeURIComponent(cleanKey)}?apiKey=${encodeURIComponent(CONFIG.API_KEY.trim())}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch project "${cleanKey}": ${response.getContentText()}`);
  }

  const result = JSON.parse(response.getContentText());
  // If an array was returned (e.g. from general list endpoint), find the matching project
  if (Array.isArray(result)) {
    const match = result.find(p => p.projectKey === cleanKey || String(p.id) === cleanKey);
    return match || result[0];
  }
  return result;
}

/**
 * Finds the issueTypeId for the given issue type name, or defaults to the first one.
 */
function getBacklogIssueTypeId(projectId, preferredName) {
  const url = `${getBaseUrl()}/api/v2/projects/${projectId}/issueTypes?apiKey=${encodeURIComponent(CONFIG.API_KEY.trim())}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch issue types: ${response.getContentText()}`);
  }

  const issueTypes = JSON.parse(response.getContentText());
  if (!issueTypes || issueTypes.length === 0) {
    throw new Error('No issue types found for this project.');
  }

  if (preferredName) {
    const match = issueTypes.find(it => it.name && it.name.toLowerCase() === preferredName.toLowerCase());
    if (match) return match.id;
  }

  // Fallback to first issue type
  return issueTypes[0].id;
}

/**
 * Resolves the numeric assignee ID from the sheet's assignee name.
 * Looks up USER_EMAIL_MAP first, then matches against project users' mailAddress or name.
 */
function resolveAssigneeId(projectId, assigneeName) {
  if (!assigneeName) return null;
  const targetEmail = USER_EMAIL_MAP[assigneeName] ? USER_EMAIL_MAP[assigneeName].toLowerCase().trim() : null;

  const url = `${getBaseUrl()}/api/v2/projects/${projectId}/users?apiKey=${encodeURIComponent(CONFIG.API_KEY.trim())}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch project users: ${response.getContentText()}`);
  }

  const users = JSON.parse(response.getContentText());

  // 1. Try matching by email from USER_EMAIL_MAP
  if (targetEmail) {
    const matchedUser = users.find(u => u.mailAddress && u.mailAddress.toLowerCase().trim() === targetEmail);
    if (matchedUser) return matchedUser.id;
  }

  // 2. Fallback: Try matching directly by user name or userId in Backlog
  const cleanName = assigneeName.toLowerCase();
  const matchedByName = users.find(u => 
    (u.name && u.name.toLowerCase() === cleanName) ||
    (u.userId && u.userId.toLowerCase() === cleanName)
  );

  return matchedByName ? matchedByName.id : null;
}

/**
 * Calls Backlog API v2 to create an issue.
 * Explicitly encodes body as application/x-www-form-urlencoded to ensure Backlog receives projectId.
 */
function createBacklogIssue(payload) {
  const url = `${getBaseUrl()}/api/v2/issues?apiKey=${encodeURIComponent(CONFIG.API_KEY.trim())}`;
  
  // Backlog strictly expects application/x-www-form-urlencoded query string format
  const formBody = Object.keys(payload)
    .filter(key => payload[key] !== null && payload[key] !== undefined && payload[key] !== '')
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(payload[key]))
    .join('&');

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: formBody,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (statusCode !== 200 && statusCode !== 201) {
    throw new Error(`Backlog API Error (${statusCode}): ${responseBody}`);
  }

  return JSON.parse(responseBody);
}

/**
 * Helper: Converts date values from Google Sheets into "YYYY-MM-DD" string.
 */
function formatDate(dateVal, timeZone) {
  if (!dateVal) return null;

  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    return Utilities.formatDate(dateVal, timeZone, 'yyyy-MM-dd');
  }

  // If entered as string, handle formats like "DD-MM-YYYY" or "YYYY-MM-DD"
  if (typeof dateVal === 'string') {
    const cleaned = dateVal.trim();
    // Match DD-MM-YYYY
    const dmy = cleaned.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (dmy) {
      const day = dmy[1].padStart(2, '0');
      const month = dmy[2].padStart(2, '0');
      const year = dmy[3];
      return `${year}-${month}-${day}`;
    }
    // Match YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(cleaned)) {
      return cleaned;
    }
  }

  return null;
}
