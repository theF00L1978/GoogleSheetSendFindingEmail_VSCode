// Version 1.04.00
// add check Saturday/Sunday

function SendReportEmail() {
  var sendOrNot = isSendMailDay();
  if (sendOrNot == 1){
    var masterSheet = SpreadsheetApp.getActiveSpreadsheet();
    var settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All_List');
    var reportName = settingSheet.getRange(2,3).getValues();
    var newSS = GetAllReportSheet(masterSheet);

    var subject = settingSheet.getRange(2,5).getValues();
    var mailTo = getRecipient(settingSheet);
    var body = settingSheet.getRange(2,6).getValues();

    var file = Drive.Files.get(newSS.getId());
    var url = file.exportLinks[MimeType.MICROSOFT_EXCEL];

    var response = UrlFetchApp.fetch(url,
      {headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()}});
    var doc = response.getBlob();
    doc.setName(getYesterdayAsString()+ reportName+ '.xlsx');

    MailApp.sendEmail(mailTo, subject, body, {attachments:[doc]});
  }
};

function GetAllReportSheet(masterSheet) {
  var settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All_List');
  var column = settingSheet.getRange('G2:G');
  var values = column.getValues();
  var ct = 0;
  var newSpreadSheetName = 'Spreadsheet to Export ' + getYesterdayAsString();
  var newSpreadSheet = SpreadsheetApp.create(newSpreadSheetName);

  SpreadsheetApp.setActiveSpreadsheet(newSpreadSheet)
  var newSpreadSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  while ( values[ct] && values[ct][0] != "" ) {
    
    var reportSheetName = values[ct][0]
    newSpreadSheet = CopyReportSheet(reportSheetName, newSpreadSheetId,masterSheet)
    ct++;
  }
  newSpreadSheet.deleteSheet(newSpreadSheet.getSheetByName('Sheet1'));
  return newSpreadSheet;

};

function CopyReportSheet(reportSheetName,newSpreadSheetId,masterSheet) {
  SpreadsheetApp.setActiveSpreadsheet(masterSheet)
  var reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(reportSheetName);

  var destination = SpreadsheetApp.openById(newSpreadSheetId);
  reportSheet.copyTo(destination);
  SpreadsheetApp.setActiveSpreadsheet(destination)
  var newSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
     
  return newSpreadSheet;
};


function getRecipient(settingSheet) {
  var column = settingSheet.getRange('D2:D');
  var values = column.getValues();
  var ct = 0;
  var allRecipient = ''
  while ( values[ct] && values[ct][0] != "" ) {
    if(ct != 0){
      allRecipient = allRecipient + ','
    }
    allRecipient = allRecipient + values[ct][0]
    ct++;
  }
  return allRecipient;
};

function getYesterdayAsString() {
  const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(getYesterday(), timezone, "yyyy.MM.dd");
};
 
function getYesterday() {
  const today = new Date();
  const yesterday = new Date(new Date().setDate(today.getDate() - 1));
  return yesterday;
};

function isSendMailDay() {
  var today = new Date();
  var dayOfWeek = today.getDay();
  var i = 0;
  if ((dayOfWeek != 0) && (dayOfWeek != 6)){
    i = 1;
  }
  return i;
}
 

