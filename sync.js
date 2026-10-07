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
 * - Can be used as a standalone Apps Script file (with config.js) OR as an Apps Script Library!
 */

// ============================================================================
// DEFAULT CONFIGURATION & LIBRARY STATE
// ============================================================================

/**
 * Default fallback configuration.
 * User-provided CONFIG in config.js or library arguments will override these values.
 */
const DEFAULT_CONFIG = {
  SPACE_DOMAIN: '',
  API_KEY: '',
  PROJECT_KEY: '',
  DEFAULT_ISSUE_TYPE: 'Task',
  PRIORITY_ID: 3,
  HEADER_ROW_INDEX: 5,
  COLUMNS: {
    TASK: 2,         // Column B (Task Summary / Issue Title)
    ASSIGNED_TO: 3,  // Column C (Assignee Name)
    START_DATE: 5,   // Column E (Start Date)
    END_DATE: 6      // Column F (End Date / Due Date)
  }
};

/**
 * In-memory active configuration and user email map (for library consumers).
 */
let _activeConfig = null;
let _activeUserEmailMap = null;

/**
 * Explicitly sets the configuration and user email map for library consumers.
 * 
 * @param {Object} config - Backlog configuration object
 * @param {Object} [userEmailMap] - Mapping of assignee names to emails
 */
function setConfig(config, userEmailMap) {
  _activeConfig = config || null;
  if (userEmailMap !== undefined) {
    _activeUserEmailMap = userEmailMap;
  }
}

/**
 * Resolves the merged configuration by combining DEFAULT_CONFIG with:
 * 1. Explicitly passed customConfig
 * 2. Pre-set _activeConfig (via setConfig)
 * 3. Globally declared CONFIG (from config.js)
 * 
 * @param {Object} [customConfig]
 * @return {Object} Merged configuration object
 */
function getEffectiveConfig_(customConfig) {
  const source = customConfig || _activeConfig || (typeof CONFIG !== 'undefined' ? CONFIG : null);
  if (!source) {
    throw new Error(
      'Configuration not found. Please provide a CONFIG object, define it in config.js, or call setConfig().'
    );
  }

  return {
    ...DEFAULT_CONFIG,
    ...source,
    COLUMNS: {
      ...DEFAULT_CONFIG.COLUMNS,
      ...(source.COLUMNS || {})
    }
  };
}

/**
 * Resolves the user email mapping from:
 * 1. Explicitly passed customUserEmailMap
 * 2. Pre-set _activeUserEmailMap (via setConfig)
 * 3. Globally declared USER_EMAIL_MAP (from config.js)
 * 
 * @param {Object} [customUserEmailMap]
 * @return {Object}
 */
function getEffectiveUserEmailMap_(customUserEmailMap) {
  if (customUserEmailMap) return customUserEmailMap;
  if (_activeUserEmailMap) return _activeUserEmailMap;
  if (typeof USER_EMAIL_MAP !== 'undefined') return USER_EMAIL_MAP;
  return {};
}

/**
 * Utility to load configuration stored in Apps Script Script Properties.
 * Useful for keeping credentials out of code entirely (File > Project Settings > Script Properties).
 * 
 * Supported properties:
 * - BACKLOG_SPACE_DOMAIN
 * - BACKLOG_API_KEY
 * - BACKLOG_PROJECT_KEY
 * - BACKLOG_DEFAULT_ISSUE_TYPE (optional)
 * - BACKLOG_PRIORITY_ID (optional)
 * - BACKLOG_HEADER_ROW_INDEX (optional)
 * - BACKLOG_USER_EMAIL_MAP (optional JSON string)
 * 
 * @return {{ config: Object, userEmailMap: Object }}
 */
function loadConfigFromScriptProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const config = {
    SPACE_DOMAIN: props.BACKLOG_SPACE_DOMAIN || '',
    API_KEY: props.BACKLOG_API_KEY || '',
    PROJECT_KEY: props.BACKLOG_PROJECT_KEY || '',
    DEFAULT_ISSUE_TYPE: props.BACKLOG_DEFAULT_ISSUE_TYPE || DEFAULT_CONFIG.DEFAULT_ISSUE_TYPE,
    PRIORITY_ID: props.BACKLOG_PRIORITY_ID ? Number(props.BACKLOG_PRIORITY_ID) : DEFAULT_CONFIG.PRIORITY_ID,
    HEADER_ROW_INDEX: props.BACKLOG_HEADER_ROW_INDEX ? Number(props.BACKLOG_HEADER_ROW_INDEX) : DEFAULT_CONFIG.HEADER_ROW_INDEX,
    COLUMNS: props.BACKLOG_COLUMNS ? JSON.parse(props.BACKLOG_COLUMNS) : DEFAULT_CONFIG.COLUMNS
  };

  let userEmailMap = {};
  if (props.BACKLOG_USER_EMAIL_MAP) {
    try {
      userEmailMap = JSON.parse(props.BACKLOG_USER_EMAIL_MAP);
    } catch (e) {
      Logger.log('Warning: Failed to parse BACKLOG_USER_EMAIL_MAP from script properties: ' + e.message);
    }
  }

  return { config, userEmailMap };
}

// ============================================================================
// MENU HELPER
// ============================================================================

/**
 * Creates the custom menu in Google Sheets.
 * Call this from your spreadsheet's onOpen() trigger.
 * 
 * @param {string} [menuTitle='Backlog Sync'] Menu title to show in Google Sheets
 * @param {string} [functionName='syncSelectedRowToBacklog'] Target function name to run on click in the client sheet
 */
function createMenu(menuTitle = 'Backlog Sync', functionName = 'syncSelectedRowToBacklog') {
  SpreadsheetApp.getUi()
    .createMenu(menuTitle)
    .addItem('Sync Selected Row', functionName)
    .addToUi();
}


// ============================================================================
// MAIN SYNC FUNCTION
// ============================================================================

/**
 * Reads the currently selected row and creates an issue in Backlog.
 * Can be called with no arguments (uses config.js) or with custom configuration.
 * 
 * @param {Object} [customConfig] - Optional configuration object
 * @param {Object} [customUserEmailMap] - Optional assignee-to-email mapping object
 */
function syncSelectedRowToBacklog(customConfig, customUserEmailMap) {
  const ui = SpreadsheetApp.getUi();

  let config;
  let userEmailMap;
  try {
    config = getEffectiveConfig_(customConfig);
    userEmailMap = getEffectiveUserEmailMap_(customUserEmailMap);

    // Validate required credentials
    if (!config.SPACE_DOMAIN || config.SPACE_DOMAIN === 'YOUR_SPACE.backlog.com') {
      throw new Error('Please configure a valid SPACE_DOMAIN in config.js or pass it to syncSelectedRowToBacklog().');
    }
    if (!config.API_KEY || config.API_KEY === 'YOUR_BACKLOG_API_KEY') {
      throw new Error('Please configure a valid API_KEY in config.js or pass it to syncSelectedRowToBacklog().');
    }
    if (!config.PROJECT_KEY || config.PROJECT_KEY === 'YOUR_PROJECT_KEY') {
      throw new Error('Please configure a valid PROJECT_KEY in config.js or pass it to syncSelectedRowToBacklog().');
    }
  } catch (err) {
    ui.alert('Configuration Error', err.message, ui.ButtonSet.OK);
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const rowIndex = activeCell.getRow();

  // 1. Validation: Prevent running on header or title rows
  if (rowIndex <= config.HEADER_ROW_INDEX) {
    ui.alert(
      'Invalid Selection',
      `Please select a valid task row below row ${config.HEADER_ROW_INDEX}.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // 2. Read row data from Columns B to F
  const maxCol = Math.max(...Object.values(config.COLUMNS));
  const rowValues = sheet.getRange(rowIndex, 1, 1, maxCol).getValues()[0];

  const taskCell = sheet.getRange(rowIndex, config.COLUMNS.TASK);
  const rawTaskText = String(rowValues[config.COLUMNS.TASK - 1] || '').trim(); // Column B
  const assigneeName = String(rowValues[config.COLUMNS.ASSIGNED_TO - 1] || '').trim(); // Column C
  const startDateRaw = rowValues[config.COLUMNS.START_DATE - 1]; // Column E
  const endDateRaw = rowValues[config.COLUMNS.END_DATE - 1]; // Column F

  // 3. Validation: Check if row is empty or a section header (e.g. "Backend Development")
  if (!rawTaskText) {
    ui.alert('Empty Task', 'The selected row does not have a task title in Column B.', ui.ButtonSet.OK);
    return;
  }

  // Extract link(s) from Column B (both plain-text https:// and rich text hyperlinks)
  const urlRegex = /https?:\/\/[^\s]+/gi;
  const plainTextUrls = (rawTaskText.match(urlRegex) || []).map(u => u.replace(/[.,;:)]+$/, ''));

  const richTextUrls = [];
  try {
    const richText = taskCell.getRichTextValue();
    if (richText) {
      const cellLink = richText.getLinkUrl();
      if (cellLink) richTextUrls.push(cellLink);
      const runs = richText.getRuns();
      if (runs) {
        runs.forEach(run => {
          const runUrl = run.getLinkUrl();
          if (runUrl) richTextUrls.push(runUrl);
        });
      }
    }
  } catch (e) {
    Logger.log('Could not read rich text link: ' + e.message);
  }

  // Deduplicate extracted URLs preserving order
  const extractedLinks = Array.from(new Set([...plainTextUrls, ...richTextUrls]));

  // Clean summary by removing URLs so issue title is neat; fallback to raw text if only URL was provided
  let summary = rawTaskText.replace(urlRegex, '').trim().replace(/\s{2,}/g, ' ');
  if (!summary) {
    summary = rawTaskText;
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
    const project = getBacklogProject(config.PROJECT_KEY, config);
    if (!project || !project.id) {
      throw new Error(`Could not find project or numeric ID for "${config.PROJECT_KEY}".`);
    }

    const issueTypeId = getBacklogIssueTypeId(project.id, config.DEFAULT_ISSUE_TYPE, config);

    // 5. Resolve Assignee ID via email mapping (proceed without assignee if not matched)
    let assigneeId = null;
    if (assigneeName) {
      assigneeId = resolveAssigneeId(project.id, assigneeName, config, userEmailMap);
      if (!assigneeId) {
        Logger.log(`Notice: Assignee "${assigneeName}" could not be matched to a Backlog user. Creating task without assignee.`);
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
      priorityId: config.PRIORITY_ID
    };

    // If a link was found in Column B, add that link alone to Backlog issue description
    if (extractedLinks.length > 0) {
      payload.description = extractedLinks.join('\n');
    }

    if (assigneeId) payload.assigneeId = assigneeId;
    if (startDate) payload.startDate = startDate;
    if (dueDate) payload.dueDate = dueDate;

    Logger.log('Payload to submit: ' + JSON.stringify(payload));

    // 8. Send POST Request to Backlog
    const createdIssue = createBacklogIssue(payload, config);
    const issueUrl = `${getBaseUrl(config)}/view/${createdIssue.issueKey}`;

    // 9. Feedback to user (does not touch or modify the sheet)
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Issue ${createdIssue.issueKey} created successfully!`,
      'Success',
      5
    );

    const descInfo = extractedLinks.length > 0 ? `\nDescription Link: ${extractedLinks.join('\n')}` : '';
    const assigneeInfo = assigneeName ? `\nAssignee: ${assigneeId ? assigneeName : 'Unassigned (not matched)'}` : '';

    ui.alert(
      'Issue Created Successfully',
      `Backlog Issue: ${createdIssue.issueKey}\nSummary: ${summary}${descInfo}${assigneeInfo}\n\nURL: ${issueUrl}`,
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
 * @param {Object} [config]
 */
function getBaseUrl(config) {
  const effectiveConfig = config || getEffectiveConfig_();
  const cleanDomain = String(effectiveConfig.SPACE_DOMAIN || '')
    .trim()
    .replace(/^https?:\/\//i, '')
    .replace(/\/+$/, '');
  return `https://${cleanDomain}`;
}

/**
 * Fetches project details from Backlog.
 * @param {string} projectKey
 * @param {Object} [config]
 */
function getBacklogProject(projectKey, config) {
  const effectiveConfig = config || getEffectiveConfig_();
  const cleanKey = String(projectKey || '').trim();
  const url = `${getBaseUrl(effectiveConfig)}/api/v2/projects/${encodeURIComponent(cleanKey)}?apiKey=${encodeURIComponent(String(effectiveConfig.API_KEY || '').trim())}`;
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
 * @param {number|string} projectId
 * @param {string} [preferredName]
 * @param {Object} [config]
 */
function getBacklogIssueTypeId(projectId, preferredName, config) {
  const effectiveConfig = config || getEffectiveConfig_();
  const url = `${getBaseUrl(effectiveConfig)}/api/v2/projects/${projectId}/issueTypes?apiKey=${encodeURIComponent(String(effectiveConfig.API_KEY || '').trim())}`;
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
 * @param {number|string} projectId
 * @param {string} assigneeName
 * @param {Object} [config]
 * @param {Object} [userEmailMap]
 */
function resolveAssigneeId(projectId, assigneeName, config, userEmailMap) {
  if (!assigneeName) return null;
  const effectiveConfig = config || getEffectiveConfig_();
  const effectiveEmailMap = userEmailMap || getEffectiveUserEmailMap_();

  const trimmedName = String(assigneeName).trim();
  const cleanName = trimmedName.toLowerCase();

  // Look up in email map (case-insensitive key lookup)
  let targetEmail = null;
  if (effectiveEmailMap[trimmedName]) {
    targetEmail = effectiveEmailMap[trimmedName].toLowerCase().trim();
  } else {
    const matchedKey = Object.keys(effectiveEmailMap).find(k => k.trim().toLowerCase() === cleanName);
    if (matchedKey) {
      targetEmail = effectiveEmailMap[matchedKey].toLowerCase().trim();
    }
  }

  const url = `${getBaseUrl(effectiveConfig)}/api/v2/projects/${projectId}/users?apiKey=${encodeURIComponent(String(effectiveConfig.API_KEY || '').trim())}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

  if (response.getResponseCode() !== 200) {
    Logger.log(`Warning: Failed to fetch project users: ${response.getContentText()}`);
    return null;
  }

  const users = JSON.parse(response.getContentText());

  // 1. Try matching by email from USER_EMAIL_MAP
  if (targetEmail) {
    const matchedUser = users.find(u => u.mailAddress && u.mailAddress.toLowerCase().trim() === targetEmail);
    if (matchedUser) return matchedUser.id;
  }

  // 2. Try exact match by Backlog user name or userId
  const matchedByName = users.find(u => 
    (u.name && u.name.trim().toLowerCase() === cleanName) ||
    (u.userId && u.userId.trim().toLowerCase() === cleanName)
  );
  if (matchedByName) return matchedByName.id;

  // 3. Fallback: match by first name or word boundary (e.g. "Shinoj" matching "Shinoj Kumar")
  const matchedPartial = users.find(u => {
    if (!u.name) return false;
    const nameParts = u.name.trim().toLowerCase().split(/\s+/);
    return nameParts.includes(cleanName) || u.name.trim().toLowerCase().startsWith(cleanName);
  });
  if (matchedPartial) return matchedPartial.id;

  return null;
}

/**
 * Calls Backlog API v2 to create an issue.
 * Explicitly encodes body as application/x-www-form-urlencoded to ensure Backlog receives projectId.
 * @param {Object} payload
 * @param {Object} [config]
 */
function createBacklogIssue(payload, config) {
  const effectiveConfig = config || getEffectiveConfig_();
  const url = `${getBaseUrl(effectiveConfig)}/api/v2/issues?apiKey=${encodeURIComponent(String(effectiveConfig.API_KEY || '').trim())}`;
  
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

// ============================================================================
// EXPORTS (CommonJS / Node environment support)
// ============================================================================
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    DEFAULT_CONFIG,
    setConfig,
    createMenu,
    syncSelectedRowToBacklog,
    getBaseUrl,
    getBacklogProject,
    getBacklogIssueTypeId,
    resolveAssigneeId,
    createBacklogIssue,
    formatDate,
    loadConfigFromScriptProperties
  };
}
