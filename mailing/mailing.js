/**
 * @OnlyCurrentDoc
*/

const EMAIL  = "Email";
const TIMESTAMP = "Delivery Timestamp";
const sheet = SpreadsheetApp.getActiveSheet();
const ui = SpreadsheetApp.getUi();
 
/**
 * onOpen - runs when sheet opens.
 * It adds 'Mailing' to the menu
 * Then checks if there are mailing scripts running and provides
 * details to the current user. Otherwise prompts to start a new email
 */
function onOpen() {
  ui.createMenu('Mailing')
      .addItem('Send Emails', 'sendEmails')
      .addItem('Clear Timestamps', 'showAlertDialog')
      .addToUi();
}

/**
 * Displays an alert dialog using the provided user interface.
 * Prompts the user if they want to send a new email.
 * If the response is 'Yes', clears timestamps.
 */
function showAlertDialog() {
  var response = ui.alert('Ready for a new adventure?',
                    'Do you want to send a new email?',
                    ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    clearTimestamps();
  }
}

/**
 * Clears timestamps in the sheet, excluding the header row.
 */
function clearTimestamps() {
  var headers = sheet.getDataRange().getValues()[0];
  var timestampIndex = headers.indexOf(TIMESTAMP) + 1; // get the column index (1-based)
  var numRows = sheet.getLastRow() - 1; // subtract one to exclude the header row
  if (timestampIndex > 0 && numRows > 0) {
    sheet.getRange(2, timestampIndex, numRows).clearContent();
    if (isTriggerPresent) {
      deleteEmailTriggers();
    }
  }
}

/**
 * Displays a custom sidebar in the Google Apps Script UI with a specified message.
 */
function displaySidebar() {
  let message = "Script still running";
  const htmlContent = `<div style="padding: 10px;"><p>${message}</p></div>`;
  const html = HtmlService.createHtmlOutput(htmlContent)
                              .setTitle('Mailing Status');
  ui.showSidebar(html);
}



/**
 * Checks whether a trigger for the 'sendEmails' function is present in the project.
 *
 * @return {boolean} True if the trigger is found, otherwise false.
 */
function isTriggerPresent() {
  var currentTrigger = ScriptApp.getProjectTriggers().find(function (trigger) {
    return trigger.getHandlerFunction() === 'sendEmails';
  });
  return !!currentTrigger;
}

/**
 * Checks whether there are any rows pending an email send action in a given sheet 
 * and deletes email triggers if all rows have been processed.
 */
function evaluateTrigger() {
  let dataRange = sheet.getDataRange();
  let data = dataRange.getValues();
  let heads = data.shift(); // Extract the headers
  const emailSentColIdx = heads.indexOf(TIMESTAMP);
  let unsentRows = data.slice(1).filter(row => row[emailSentColIdx] === '').length;
  
  if (unsentRows === 0) {
    deleteEmailTriggers();
  }
}

/**
 * Deletes email triggers associated with the 'sendEmails' handler function.
 */
function deleteEmailTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Creates an email trigger for the 'sendEmails' function
 * if no existing trigger is present.
 */
function createEmailTrigger() {
  if (!isTriggerPresent()) {
    ScriptApp.newTrigger('sendEmails')
      .timeBased()
      .everyHours(1)
      .create();
  }
}

/**
 * Prompts the user to input or copy/paste the subject line of the Gmail draft message.
 * Returns the subject line or null if canceled or left blank.
 *
 * @return {string|null} The subject line or null.
 */
function promptForSubjectLine() {
  const text = "Type or copy/paste the subject line of the Gmail draft message you would like to send:";
  let subjectLine = Browser.inputBox("Send Emails", text, Browser.Buttons.OK_CANCEL);
  if (subjectLine === "cancel" || subjectLine === ""){ 
    return null;
  }
  return subjectLine;
}

/**
 * Sends an email using the provided email template and data from the specified row.
 *
 * @param {Array} row - The data row to extract information from.
 * @param {Object} emailTemplate - The email template containing message, inline images, and attachments.
 * @param {Array} columnNames - The column names corresponding to the data in the row.
 */
function sender(row, emailTemplate, columnNames) {
  let emailData = {};
  columnNames.forEach((name, index) => {
    emailData[name] = row[index];
  });

  let msgObj = fillInTemplateFromObject_(emailTemplate.message, emailData);
  GmailApp.sendEmail(row[columnNames.indexOf(EMAIL)], msgObj.subject, msgObj.text, {
    htmlBody: msgObj.html,
    // bcc: 'a.bbc@email.com, b.bbc@email.com',
    // cc: 'a.cc@email.com, b.cc@email.com',
    from: 'hi@stylebitt.com',
    name: 'Stylebitt',
    replyTo: 'hi@stylebitt.com',
    // noReply: true,
    inlineImages: emailTemplate.inlineImages,
    attachments: emailTemplate.attachments
  });
}
 
/**
 * Sends emails from sheet data.
*/
function sendEmails() {
  if (isTriggerPresent()) {
    displaySidebar();
    return;
  }
  let subjectLine = promptForSubjectLine();
  if (!subjectLine) return;
  
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  let dataRange = sheet.getDataRange();
  let data = dataRange.getValues();
  let heads = data.shift(); 
  const emailSentColIdx = heads.indexOf(TIMESTAMP);
  const batchSize = 100;

  if (data.length > batchSize) {
    createEmailTrigger();
    console.log('Trigger Created!')
  }

  let rowsToProcess = data.filter(row => row[emailSentColIdx] === '');
  if (rowsToProcess.length > 0) {
    let rowsToSend = rowsToProcess.slice(0, batchSize);
    let out = rowsToSend.map((row, index) => {
      try {
        sender(row, emailTemplate, heads);
        return [new Date()];
      } catch(e) {
        return [e.message]
      }
    })

    let startRow = dataRange.getRow() + data.indexOf(rowsToSend[0]);
    sheet.getRange(startRow + 1, emailSentColIdx + 1, out.length).setValues(out);
    evaluateTrigger();
  }
}
  
/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
*/
function getGmailTemplateFromDrafts_(subject_line){
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    Logger.log(draft);
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody(); 

    // Creates an inline image object with the image name as key 
    // (can't rely on image index as array based on insert order)
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

    //Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    //Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
            attachments: attachments, inlineImages: inlineImagesObj };
  } catch(e) {
    throw new Error("Oops - can't find Gmail draft");
  }

  /**
   * Filter draft objects with the matching subject linemessage by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} GmailDraft object
  */
  function subjectFilter_(subject_line){
    return function(element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
*/
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return  JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
*/
function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
};

