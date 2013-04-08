/**
 * Google Apps Mail Merge script.
 * @author: Alex Skrypnyk (alex.designworks@gmail.com)
 * @version: 1.2
 * License: GPL v2+ (http://www.gnu.org/licenses/old-licenses/gpl-2.0.html)
 *
 * @description
 * Google Apps mail Merge script that uses Gmail Drafts as a source of template.
 * Allows using attachments and inline images.
 * Also supports contact information importing from Contacts groups.
 *
 * 1. Write your email and save it as Draft.
 *    You may attach images and attachments as you would normally do.
 * 2. The words that needs to be replaced for each recipient are called
 *    "placeholders".  Replace all placeholders with unique names, surrounded
 *    by %% and %%.  You may use any characters, including spaces, in your
 *    placeholder name.
 *    Example: Hello %%First Name%%.
 * 3. Create a separate column for each placeholder in the spreadsheet and fill
 *    it with values.
 *    First cell of each column (column header) must be the same as you have
 *    specified in the body of your email.
 *    Example:  for placeholder %%First Name%%, column name will be First Name
 * 4. Run mail merge.
 *
 * To re-send the message to selected recipients, clear cell in 'Sent status'
 * column for this recipient.
 *
 * It is a good practice to create separate spreadsheet for each mail merge to
 * be able to track all sent correspondence.
 *
 * Video Tutorial
 * @see http://www.youtube.com/watch?v=WWb3hpXLrag
 *
 * Project page on GitHub
 * https://github.com/alexdesignworks/MailMerge-Done-Right
 *
 */

var placeholder_start = '%%';
var placeholder_finish = '%%';
var version = '1.2';

/**
 * Create menu items.
 */
function onInstall() {
  onOpen();
}

/**
 * Create menu items.
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Mail Merge", [
    {name: "Help", functionName: "help"},
    {name: "Import contacts from Group", functionName: "importFromGroup"},
    {name: "Start Mail Merge", functionName: "startMailMerge"}
  ]);
}
/**
 * How to use callback.
 */
function help() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setWidth(600).setHeight(480);
  var html = '';
  html += '<h1>Mail Merge Done Right!</h1>';
  html += '<p><strong>Version:</strong> ' + version + '</p>';
  html += '<ol>';
  html += '  <li>Write your email and save it as Draft.<br />';
  html += '  You may attach images and attachments as  you would normally do.</li>';
  html += '  <li>The words that needs to be replaced for each recipient  are called &quot;placeholders&quot;.  Replace all placeholders with unique names,  surrounded by ' + placeholder_start + ' and ' + placeholder_finish + '.  You may use any  characters, including spaces, in your placeholder name.<br />';
  html += '  Example: Hello ' + placeholder_start + 'First Name' + placeholder_finish + '.</li>';
  html += '  <li>Create a separate column for each placeholder in  the spreadsheet and fill it with values. <br />';
  html += '  First cell of each column (column header) must be the same as you have  specified in the body of your email.<br />';
  html += '  Example:  for placeholder ' + placeholder_start + 'First Name' + placeholder_finish + ',  column name will be <strong>First Name</strong></li>';
  html += '  <li>Run mail merge.</li>';
  html += ' </ol>';
  html += '<p>To re-send the message to selected recipients, clear cell in "Sent status" column for this recipient.</p>';
  html += '<p>It is a good practice to create separate spreadsheet for  each mail merge to be able to track all sent correspondence.</p>';
  html += '<h3>Video tutorial</h3>';
  app.add(app.createHTML(html));
  app.add(app.createAnchor('Click here to see video tutorial', 'http://www.youtube.com/watch?v=WWb3hpXLrag'));
  html = '<h3>Project page and bug reports</h3>';
  app.add(app.createHTML(html));
  app.add(app.createAnchor('MailMerge Done Right on GitHub', 'https://github.com/alexdesignworks/MailMerge-Done-Right'));
  doc.show(app);
}

/**
 * Import contacts from group.
 */
function importFromGroup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Select Group').setWidth('250').setHeight('80');

  // Fetch groups.
  var groups = ContactsApp.getContactGroups();
  var listBox = app.createListBox().setName('groups').addItem('Select...');

  for (i in groups) {
    listBox.addItem(groups[i].getName());
  }

  // Create and assign onChange handler to get selected index.
  var handler = app.createServerChangeHandler('getSelectedGroupsItem').addCallbackElement(listBox);
  listBox.addChangeHandler(handler);

  // Create run button.
  var buttonRun = app.createButton("Import contacts").setId("buttonPopulateContacts");
  buttonRun.addClickHandler(app.createServerChangeHandler('populateContacts'));

  // Create cancel button.
  var buttonCancel = app.createButton("Cancel").setId("buttonCancel");
  buttonCancel.addClickHandler(app.createServerChangeHandler('closePanel'));

  // Create panel.
  var panel = app.createVerticalPanel().setId('panelGroup');
  panel.add(app.createLabel('Select the group from your contacts'));
  panel.add(listBox);
  var buttonsPanel = app.createHorizontalPanel().setId('panelGroupButtons');
  buttonsPanel.add(buttonRun);
  buttonsPanel.add(buttonCancel);
  panel.add(buttonsPanel);
  // Add panel to app.
  app.add(panel);
  // Show window.
  doc.show(app);
}

/**
 * Start mail merge UI callback.
 */
function startMailMerge() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  if (isValuesEmpty(doc.getActiveSheet().getDataRange().getValues())){
    Browser.msgBox("No data provided in the spreadsheet.");
    return;
  }

  // Create window with title.
  var app = UiApp.createApplication().setWidth('300').setHeight('80').setTitle('Mail Merge');

  // Fetch drafts from gmail account.
  var templates = GmailApp.search("in:drafts");
  // Fill-in listbox with items.
  var listBox = app.createListBox().setName('templates').addItem('Select...');
  for (i in templates) {
    listBox.addItem((parseInt(i) + 1) + '- ' + templates[i].getFirstMessageSubject().substr(0, 40));
  }

  // Create and assign onChange handler to get selected index.
  var ListBoxHandler = app.createServerChangeHandler('getSelectedTemplatesItem').addCallbackElement(listBox);
  listBox.addChangeHandler(ListBoxHandler);

  // Create run button.
  var buttonRun = app.createButton("Run Mail Merge").setId("buttonRun");
  buttonRun.addClickHandler(app.createServerChangeHandler('sendEmails'));

  // Create cancel button.
  var buttonCancel = app.createButton("Cancel").setId("buttonCancel");
  buttonCancel.addClickHandler(app.createServerChangeHandler('cancelSend'));

  // Create panel.
  var panel = app.createVerticalPanel().setId('panel');
  // Add all UI components to the panel.
  panel.add(app.createLabel('Select the template (from your drafts in Gmail)'));
  panel.add(listBox);
  var buttonsPanel = app.createHorizontalPanel().setId('panelMailmergeButtons');
  buttonsPanel.add(buttonRun);
  buttonsPanel.add(buttonCancel);
  panel.add(buttonsPanel);
  // Add panel to app.
  app.add(panel);
  // Show window.
  doc.show(app);
}

/**
 * Cancel button callback.
 * Close app window and show cancellation toast.
 */
function cancelSend() {
  var app = UiApp.getActiveApplication();
  app.close();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Mail Merge was cancelled. No messages were sent.', 'Mail Merge', 5);
  return app;
}

/**
 * Cancel button callback.
 * Close app window and show cancellation toast.
 */
function cancelSend() {
  var app = UiApp.getActiveApplication();
  app.close();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Mail Merge was cancelled. No messages were sent.', 'Mail Merge', 5);
  return app;
}

/**
 * Close panel button callback.
 */
function closePanel() {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}


/**
 * Listbox onChange callback for Templates.
 * Get selected Templates ListBox item.
 */
function getSelectedTemplatesItem(e) {
  // 'templates' in e.parameter.templates is the name of listbox.
  ScriptProperties.setProperty("selectedTemplate", e.parameter.templates);
}

/**
 * Listbox onChange callback for Groups.
 * Get selected Groups ListBox item.
 */
function getSelectedGroupsItem(e) {
  // 'groups' in e.parameter.group is the name of listbox.
  ScriptProperties.setProperty("selectedGroup", e.parameter.groups);
}

/**
 * Contact fetching callback.
 */
function populateContacts(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getActiveSheet();
  var groupName = ScriptProperties.getProperty("selectedGroup");
  if (groupName == 'Select...') {
    // Selection has not been made.
    return;
  }

  // Ask for column names.
  var firstsnameColumn = Browser.inputBox("Which column contains First Name ? (A, B, C, ...)");
  if (firstsnameColumn == 'cancel') {
    return;
  }

  var lastnameColumn = Browser.inputBox("Which column contains Last Name ? (A, B, C, ...)");
  if (lastnameColumn == 'cancel') {
    return;
  }

  var emailColumn = Browser.inputBox("Which column contains Email ? (A, B, C, ...)");
  if (emailColumn == 'cancel') {
    return;
  }

  var rangeA1 =  dataSheet.getDataRange().getA1Notation();
  var lastColumnName =rangeA1.split(':')[0].replace(/[0-9]+/g, '');
  ss.toast(lastColumnName);

  // Check that columns names exist and add if they are not.
  columnNameRowNumber = 1;
  if (dataSheet.getRange(firstsnameColumn + columnNameRowNumber).getValue() == ""){
    dataSheet.getRange(firstsnameColumn + columnNameRowNumber).setValue("First Name");
  }

  if (dataSheet.getRange(lastnameColumn+ columnNameRowNumber).getValue() == ""){
    dataSheet.getRange(lastnameColumn + columnNameRowNumber).setValue("Last Name");
  }

  if (dataSheet.getRange(emailColumn+ columnNameRowNumber).getValue() == ""){
    dataSheet.getRange(emailColumn + columnNameRowNumber).setValue("Email Address");
  }


  // Get contacts from Contacts.
  var contacts = ContactsApp.getContactGroup(groupName).getContacts();

  for (var i in contacts) {
    var emails = contacts[i].getEmails();
    if (typeof emails === 'undefined' || emails.length === 0){
      // Skip contacts with no emails.
      continue;
    }
    var email = emails[0].getAddress();
    if (email) {
      var newRowNumber = dataSheet.getLastRow() + 1;
      dataSheet.getRange(firstsnameColumn + newRowNumber).setValue(contacts[i].getGivenName());
      dataSheet.getRange(lastnameColumn + newRowNumber).setValue(contacts[i].getFamilyName());
      dataSheet.getRange(emailColumn + newRowNumber).setValue(email);
    }
  }
}

/**
 * Send emails callback.
 * The core.
 */
function sendEmails(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var selectedTemplate = ScriptProperties.getProperty("selectedTemplate");
  if (selectedTemplate == 'Select...') {
    // Selection has not been made.
    return;
  }

  var dataSheet = ss.getActiveSheet();
  var headersRange = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn());

  // Add 2 more columns if they are not already set.
  var cellSentStatus = searchRange(dataSheet, 'Sent status', headersRange);
  if (!cellSentStatus) {
    cellSentStatus = dataSheet.getRange(1, dataSheet.getLastColumn() + 1, 1, 1);
    cellSentStatus.setValue('Sent status');
  }
  var cellSentTimestamp = searchRange(dataSheet, 'Sent timestamp', headersRange);
  if (!cellSentTimestamp) {
    cellSentTimestamp = dataSheet.getRange(1, dataSheet.getLastColumn() + 1, 1, 1);
    cellSentTimestamp.setValue('Sent timestamp');
  }

  // Searching for email column.
  var headers = headersRange.getValues();
  var emailColumnFound = false;
  for (i in headers[0]) {
    if (headers[0][i] == "Email Address") {
      emailColumnFound = true;
    }
  }
  if (!emailColumnFound) {
    // If no column found - ask for it.
    var emailColumn = Browser.inputBox("Which column contains emails of recipients ? (A, B,...)");
    if (emailColumn == 'cancel') {
      // No email - no mail merge - quit.
      return;
    }
    // Set header of emails column.
    dataSheet.getRange(emailColumn + '' + 1).setValue("Email Address");
  }
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());

  // Find template in Gmail drafts.
  var foundTemplate = GmailApp.search("in:drafts")[(parseInt(selectedTemplate.substr(0, 2)) - 1)].getMessages()[0];
  // Template body.
  var templateBody = foundTemplate.getBody();
  // Get all template attachments, including inline images.
  var templateAttachments = foundTemplate.getAttachments();
  // Notify user that sending is in progress.
  ss.toast('Sending messages.Do not make any changes to spreadsheet until complete.', 'Mail Merge', -1);

  // Searching for inline images.
  // Known issue: The code below will work only if added inline images were not
  // removed from the body. If at least 1 inline image gets removed - all inline
  // images will be randomized. This happens due to templateAttachments holding
  // all inline images, including removed ones.
  // TODO: identify removed images and clean up templateAttachments

  // Create template reg exp to locate all images within the body.
  var templateRegExp = new RegExp(foundTemplate.getId(), "g");
  if (templateBody.match(templateRegExp) != null) {
    // Get inline images count.
    var imgCount = templateBody.match(templateRegExp).length;
    var imgTags = templateBody.match(/<img[^>]+>/g);
    var imgToReplace = [];
    for (var i = 0; i < imgTags.length; i++) {
      if (imgTags[i].search(templateRegExp) != -1) {
        var imgId = imgTags[i].match(/Inline\simage[s]?\s(\d)/);
        imgToReplace.push([parseInt(imgId[1]), imgTags[i]]);
      }
    }
    // Sort array of images to be replaced.
    imgToReplace.sort(function (a, b) {
      return a[0] - b[0];
    });

    var inlineImages = {};
    // Replacing images and removing attachments.
    for (var i = 0; i < imgToReplace.length; i++) {
      // Get current attachment id.
      var attId = i + (templateAttachments.length - imgCount);
      // Provide replacement title.
      var title = 'InlineImages' + i;
      inlineImages[title] = templateAttachments[attId].copyBlob().setName(title);
      // Remove inline image attachment from attachments list.
      templateAttachments.splice(attId, 1);
      // Create new img tag  - replace src attribute and value for image.
      var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + title + "\"");
      // Replace whole img tag with a new one.
      templateBody = templateBody.replace(imgToReplace[i][1], newImg);
    }
  }

  // Create one JavaScript object per row of data.
  var objects = getRowsData(dataSheet, dataRange);

  var sentEmails = 0;
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];
    // Send only emails that do not have "Email sent" in the 'Sent status' column.
    if (rowData.sentStatus != "Email sent") {
      // Generate a personalized email.
      var emailText = fillInTemplateFromObject(templateBody, rowData);
      var emailSubject = foundTemplate.getSubject();
      // Send email.
      GmailApp.sendEmail(rowData.emailAddress, emailSubject, emailText, {
        attachments: templateAttachments,
        htmlBody: emailText,
        inlineImages: inlineImages,
        from: foundTemplate.getFrom().match(/[^ <>]+@[^ <>]+/)[0]
      });
      // Fill-in 'Email status' and 'Sent timestamp'.
      dataSheet.getRange(i + 2, dataSheet.getLastColumn() - 1).setValue("Email sent");
      dataSheet.getRange(i + 2, dataSheet.getLastColumn()).setValue(new Date().toString());
      sentEmails++;
    }
  }
  if (sentEmails > 0) {
    ss.toast(sentEmails + ' of ' + objects.length + ' emails were sent', 'Mail Merge', 10);
  }
  else {
    ss.toast('None of ' + objects.length + ' emails were sent as they were sent before.Remove "Email sent" from "Sent status" column to resend.', 'Mail Merge', 10);
  }
  app.close();
  return app;
}

/**
 * Replace tokens in the template with values.
 * @param {String} template
 *   String containing tokens. Tokens are %column_names%. Tokens without data provided gets removed.
 * @param {Object} data
 *   Object with column names as sanitized properties.
 * @return {String}
 *   Tokenized template body.
 */
function fillInTemplateFromObject(template, data) {
  var body = template;
  var pattern=regExpQuote(placeholder_start)+"[^"+regExpQuote(placeholder_start)+"]+"+regExpQuote(placeholder_finish);
  var tokens = template.match(new RegExp(pattern, "ig"));
  if (tokens != null) {
    // Replace variables from the template with the actual values from the data object.
    // If no value is available, replace with the empty string.
    for (var i = 0; i < tokens.length; ++i) {
      var tokenValue = data[normalizeHeader(tokens[i])];
      body = body.replace(tokens[i], tokenValue || "");
    }
  }
  return body;
}

//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
// @see https://developers.google.com/apps-script/articles/reading_spreadsheet_data
//
//////////////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    }
    else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

/**
 * Check that provided data is empty or trimmed-empty.
 * @param cellData
 * @return {Boolean}
 */
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && (cellData == "" || cellData.replace(/^\s+|\s+$/g,'') == "");
}

/**
 * Checks that all values in provided array are empty.
 * @param values
 */
function isValuesEmpty(values){
  var empty = true;
  for (var i in values){
    if (typeof values[i] === 'array' || typeof values[i] === 'object'){
      empty = isValuesEmpty(values[i]);
    }
    else{
      if (!isCellEmpty(values[i])){
        empty = false;
      }
    }
    // If at least one is not empty - quit.
    if (!empty){
      return empty;
    }
  }
  return empty;
}


/**
 * Returns true if the character char is alphabetical, false otherwise.
 */
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

/**
 * Returns true if the character char is a digit, false otherwise.
 */
function isDigit(char) {
  return char >= '0' && char <= '9';
}


/**
 * Search for value in the range and return range of first occurence
 */
function searchRange(sheet, needle, range) {
  if (!range) {
    range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  }

  var leftColumn = range.getColumnIndex();
  var rightColumn = range.getLastColumn();
  var topRow = range.getRowIndex();
  var bottomRow = range.getLastRow();

  for (var i = topRow; i <= bottomRow; i++) {
    for (var j = leftColumn; j <= rightColumn; j++) {
      if (needle == sheet.getRange(i, j, 1, 1).getValue()) {
        return sheet.getRange(i, j);
      }
    }
  }

  return false;
}
/**
 * Array to object with empty elements filtering.
 */
function toObject(arr) {
  var rv = {};
  for (var i = 0; i < arr.length; ++i) {
    if (arr[i] !== undefined) {
      rv[i] = arr[i];
    }
  }
  return rv;
}
/**
 * Object to array with empty elements filtering.
 */
function toArray(obj) {
  var arr = [];
  for (i in obj) {
    if (obj[i] != null) {
      arr.push(obj[i]);
    }
  }
  return arr;
}
/**
 * Escape character for regexp.
 */
function regExpQuote (str) {
  return str.replace(/([.?*+^$[\]\\(){}|-])/g, "\\$1");
};
