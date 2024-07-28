/**
 * @OnlyCurrentDoc Limits the script to only accessing the current document.
 */

var SIDEBAR_TITLE = 'Marketing Campaign Planner';

/**
 * Adds a custom menu that opens the sidebar directly in Google Sheets.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Open Marketing Campaign Planner', 'showSidebar')
    .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initialization work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle(SIDEBAR_TITLE)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Adds a marketing campaign event to Google Calendar.
 *
 * @param {Object} campaign Details of the marketing campaign.
 */
function addCampaignToCalendar(campaign) {
  var startDateTime = new Date(campaign.startDate + 'T' + formatTime(campaign.startTime));
  var endDateTime = new Date(campaign.endDate + 'T' + formatTime(campaign.endTime));
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var calendar = CalendarApp.getDefaultCalendar();
  calendar.createEvent(campaign.title, startDateTime, endDateTime, {
    description: campaign.description,
    timeZone: timeZone
  });
}

/**
 * Sends an email campaign using Gmail.
 *
 * @param {Object} campaign Details of the email campaign.
 */
function sendEmailCampaign(campaign) {
  MailApp.sendEmail({
    to: campaign.recipients,
    subject: campaign.subject,
    htmlBody: campaign.body
  });
}

/**
 * Retrieves the active cell value and converts it to the appropriate type.
 *
 * @param {String} field The field to pull the value for.
 * @return {String} The value of the active cell.
 */
function getActiveValue(field) {
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  var value = cell.getValue();

  if (field.includes('date')) {
    return Utilities.formatDate(new Date(value), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  } else if (field.includes('time')) {
    return Utilities.formatDate(new Date(value), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm');
  } else {
    return value;
  }
}

/**
 * Converts time from `HH:mm` format to `HH:mm:ss` format for calendar event creation.
 *
 * @param {String} time The time in `HH:mm` format.
 * @return {String} The time in `HH:mm:ss` format.
 */
function formatTime(time) {
  return time + ':00';
}

/**
 * Executes specified action for managing campaign data.
 *
 * @param {String} action An identifier for the action to take.
 * @param {Object} campaign Details of the campaign.
 */
function manageCampaigns(action, campaign) {
  if (action == "addEvent") {
    addCampaignToCalendar(campaign);
  } else if (action == "sendEmail") {
    sendEmailCampaign(campaign);
  }
}

/**
 * Sets up triggers to ensure the addon menu appears in Sheets.
 */
function setupTriggers() {
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();
  
  ScriptApp.newTrigger('onInstall')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onInstall()
    .create();
}

/**
 * Deletes all triggers for this project.
 */
function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}