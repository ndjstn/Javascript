/*
  - This system will send an email campain based upon value in a column using Google sheets.
  - This is considered a "bound" script, meaning it is intristically tied to a spreadsheet.
  - The document's used will send their title as the subject line.
  - The document's body will be the message.
  - This requires that the sheet has 3 seperate column names which are: Email, Campaign, and Name. 
    These column names will be case sensitive.
  Erratta:
  Bear in mind that Google has daily email limits that can come out of different types of email addresses. You should check with them.
  You should also consider an application that attempts to rewrite both the subject line and content of the body of the message as things 
  may inadvertently be picked up by a spam filter. Another thought would be to wright or use something else to scrape data from sites to 
  use as a source for target data. 
*/

// Send the user and the campain email to the function
function emailer(recipient, document_to_email){ // This takes an email and a google doc id to send.
  
  // The recipient 
  const email = recipient;
  Logger.log("Sending email to: " + email);
  
  // The document to send
  const doc = DocumentApp.openById(document_to_email); // It sends so I am not sure I care about this
  const body = doc.getBody();
  let email_content = body.editAsText().getText(); 
  
  // Email Subject Line
  Logger.log("With Subject Line: " + doc.getName());
  const subject = doc.getName();
  
  // Send the mail
  MailApp.sendEmail(email, subject, email_content);
  // Log the email sent or code if given
  Logger.log("Email Content: " + email_content); 
  
}

// Get the name of the column
function getByName(colName, row) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    return data[row-1][col];
  }
}

function main() {
  // Campaign sheets
  // TODO: Expand to either iter over a list
  // TODO: Expand to take in Docnames
  const campaign_1 = ''; // Place doc id here
  const campaign_2 = ''; // Place doc id here
  const campaign_3 = '' ; // Place doc id here

  // Bounding for iterables
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const number_of_cols = parseInt(sheet.getLastColumn());
  const number_of_rows = parseInt(sheet.getLastRow());
  
  // Itter over rows and do work
  // TODO: Remove case sensitivity from column names.
  for (var i = 1; i < number_of_rows; i++) {
    // Pointers Needed to email
    // TODO: add a function to check if email is in a do not contact list before sending.
    receiver_email = getByName('Email',i); 
    // Pointers Needed to campain
    campaign_message = getByName('Campaign',i);
    // Greeting for for the recipient 
    // TODO: could be split into first/last/etc
    greeting = getByName('Name',i);

    Logger.log("This is row: "+ (i+1));
    // Debug msg
    Logger.log("Emailing to: " + receiver_email + " With Campaign: " + campaign_message + " sent by: " + greeting);
    
    // Message selection logic ...  switch to case for speed, or possibly a single matching if?
    if (campaign_message == 1) {
      emailer(receiver_email, campaign_1);
    } else if (campaign_message == 2) {
        emailer(receiver_email, campaign_2);
    } else if (campaign_message == 3) {
        emailer(receiver_email, campaign_3);
    } else {
      // Log an error message
      Logger.log("Issue with campain " + (campaign_message) + " to " + receiver_email)
    }
  }
} 

