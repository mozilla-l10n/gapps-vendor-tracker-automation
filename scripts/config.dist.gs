// Configuration template. Copy this file to config.gs (which is gitignored)
// and fill in the values for your deployment. Apps Script shares one global
// scope across all .gs files, so these constants are visible everywhere.
//
// Secrets do NOT go here. The Jira email and API token live in the Apps
// Script project under Project Settings > Script properties, as JIRA_EMAIL
// and JIRA_TOKEN. Never put them in a file.

// Jira Cloud base URL, e.g. https://<site>.atlassian.net (no trailing slash).
const JIRA_BASE = '';

// Key of the Jira project used for vendor intake tickets, e.g. ABC.
const JIRA_PROJECT = '';

// Jira custom field IDs from the intake form, e.g. customfield_12345.
// To find an ID: GET <JIRA_BASE>/rest/api/3/field and search by field name.
const JIRA_CC_FIELD = '';        // cost center, e.g. "CC12345" or "12345"
const JIRA_NAME_FIELD = '';      // tracker row name
const JIRA_DEADLINE_FIELD = '';  // deadline (yyyy-mm-dd)
const JIRA_LOCALE_FIELD = '';    // target locales, e.g. "tr, it"
const JIRA_INVOICED_FIELD = '';  // invoiced amount (number)

// Drive folder ID of the invoice folder for INVOICE_YEAR (the trailing
// segment of the folder URL, .../folders/<ID>). Month subfolders inside it
// must be named like "07-July".
const INVOICE_FOLDER_ID = '';

// Calendar year the invoice folder covers, e.g. '2026'. Update this and
// INVOICE_FOLDER_ID together at the start of each year.
const INVOICE_YEAR = '';

// Jira assignee first name -> name used in the tracker's Mozilla PM column,
// e.g. { 'Firstname': 'Nickname' }. Leave {} to always use the Jira name.
const PM_ALIAS = {};
