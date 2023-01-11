// import the SpreadsheetApp and MailApp classes
var SpreadsheetApp = SpreadsheetApp;
var MailApp = MailApp;

// import the DriveApp class
var DriveApp = DriveApp;

// import the MimeType enum
var MimeType = DriveApp.MimeType;

// import the DocumentApp class
var DocumentApp = DocumentApp;

// function to convert text values in the Google Sheet into numeric values
function convertToNumeric() {
  // retrieve the active sheet
  var sheet = SpreadsheetApp.getActiveSheet();

  // get the data range of the sheet
  var data = sheet.getDataRange();

  // get the values of the data range as a 2D array
  var values = data.getValues();

  // loop through the rows of the array
  for (var i = 0; i < values.length; i++) {
    // loop through the columns of the array
    for (var j = 0; j < values[i].length; j++) {
      // check if the current cell contains a text value that needs to be converted
      if (values[i][j] == "Low") {
        // convert the text value to a numeric value
        values[i][j] = 1;
      } else if (values[i][j] == "Medium") {
        values[i][j] = 2;
      } else if (values[i][j] == "High") {
                values[i][j] = 3;
      }
    }
  }
  // set the values of the data range to the updated 2D array
  data.setValues(values);
}

function autoFillIPSTemplateGoogleDoc(e) {
  // call the convertToNumeric function
  convertToNumeric();

  // declare variables from Google Sheet
  let investorName = e.values[1];
  let timeStamp = e.values[0];
  let emailID = e.values[2];

  //retrive the values from the sheet after conversion
  var sheet = SpreadsheetApp.getActiveSheet();
  var riskTolerance = sheet.getRange(2, 3).getValue();
  var timeHorizon = sheet.getRange(2, 4).getValue();

  // calculate score based on risk tolerance
  let riskScore = 0;
  if (riskTolerance >= 0 && riskTolerance <= 5) {
    riskScore = 1;
  } else if (riskTolerance > 5 && riskTolerance <= 10) {
    riskScore = 2;
  } else {
    riskScore = 3;
  }

  // calculate score based on time horizon
  let timeScore = 0;
  if (timeHorizon >= 0 && timeHorizon <= 5) {
    timeScore = 1;
  } else if (timeHorizon > 5 && timeHorizon <= 10) {
    timeScore = 2;
  } else {
    timeScore = 3;
  }

  // declare goal variables
  let goal1 = ""
  let goal2 = ""
  let goal3 = ""

  // convert value from column 5 of Google Sheet to string
  const goals = e.values[5].toString();

  //create an array and parse values from CSV format, store them in an array
  goalsArr = goals.split(',')
  if (goalsArr.length >= 1)
    goal1 = goalsArr[0]
  if (goalsArr.length >= 2)
    goal2 = goalsArr[1]
  if (goalsArr.length >= 3)
    goal3 = goalsArr[2]

  //grab the template file ID to modify
  const file = DriveApp.getFileById(templateID);

  //grab the Google Drive folder ID to place the modied file into
  var folder = DriveApp.getFolderById(folderID)

  //create a copy of the template file to modify, save using the naming conventions below
  var copy = file.makeCopy(investorName + ' Investment Policy', folder);

  //modify the Google Drive file
  var doc = DocumentApp.openById(copy.getId());

  var body = doc.getBody();

  body.replaceText('%InvestorName%', investorName);
  body.replaceText('%Date%', timeStamp);
  body.replaceText('%RiskScore%', riskScore);
  body.replaceText('%TimeScore%', timeScore);
  body.replaceText('%Goal1%', goal1.trim())
  body.replaceText('%Goal2%', goal2.trim())
  body.replaceText('%Goal3%', goal3.trim())

  doc.saveAndClose();

  //find the file that was just modified, convert to PDF, attach to e-mail, send e-mail
  var attach = DriveApp.getFileById(copy.getId());
  var pdfattach = attach.getAs(MimeType.PDF);
  MailApp.sendEmail(emailID, subject, emailBody, { attachments: [pdfattach] });
}
