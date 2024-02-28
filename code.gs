//Automagically send emails using Google Sheets – September 05, 2022
//

function sendWish() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MainSheet");
  var startRow = 2; // First row of data to process
  var numRows = sheet.getLastRow() - 1; // Number of rows to process
  var dataRange = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()); // Fetch the range of cells being used A2:LastUsed
  var data = dataRange.getValues(); // Fetch values for each row in the Range.

  //!!!!!!!!!!!!!!!!!!!!!!!!!var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(1,1).getValue(); //Get template text from first cell in Template sheet
  var EMAIL_SENT = 'EMAIL_SENT';

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var date = new Date();
    var sheetDate = new Date(row[2]);

    //Make date formats the same for comparisson
    Sdate = Utilities.formatDate(date, 'GMT+6', 'EEE, MMM d, yyyy')
    SsheetDate = Utilities.formatDate(sheetDate, 'GMT+6', 'EEE, MMM d, yyyy')

    //Look through sheet for date in the first row that corresponds to today to build emails
    if (Sdate == SsheetDate) {
      if (row[3] != EMAIL_SENT) {
        var emailAddress = row[1];
        var subject = "Happy Birthday";

        var Name = row[0];
        //var SundayDate = row[3];
        //var Assignment = row[4];
        //var Theme = row[5];

        //var SundayDate = Utilities.formatDate(SundayDate,’GMT-0300′,’MMM d’);

        // Another option for the email text is to concatinate it directly rather than use a template with replacement
        //var emailText = "Happy Birthday to, " + row[0];
        var emailText=HtmlService.createHtmlOutputFromFile('Body').getContent();

        //Replace the {STANDINS} in the template with the values assigned to variables from the Data Spreadsheet
        //var emailText = templateText.replace("{SundayDate}", SundayDate).replace("{ChildName}",ChildName).replace("{Assignment}",Assignment).replace("{Theme}",Theme);

        //Use the logger here to check that your template replacement has worked properly
        //Logger.log(emailText);

        //Send the mail – Once everything is set up properly, this next line is what sends the emails
        MailApp.sendEmail({
          name: 'Azraf Sami',
          to: row[1], 
          subject: "Happy Birthday", 
          htmlBody: emailText});

        //Add Email sent indication to end of row
        sheet.getRange(startRow + i, 4).setValue("EMAIL_SENT");

        //Email yourself to let you know what email was sent
        //Logger.log('SENT : ' + emailAddress + ' ' + subject + ' ' + emailText)
        //var body = Logger.getLog();

      }
    }
  }
}
