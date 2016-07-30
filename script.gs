var SPREADSHEET = '1j2xLQ8Z1sa0ny2mp77Ipu1-LEDgyuubUWFhZalQ9Dqk';
var INFO_FLYER_YES = 'ja';
var SHEET_TEAMS = 'Fragebogen - Teams';
var INFO_FLYER = 'Info Flyer';
var EMAIL_ADDRESS_FIELD = 'E-mail address';

var EMAIL_ADDRESS = 'E-mail address';
var EMAIL_REPLY_TO_ADDRESS = 'info@iha.help';
var EMAIL_DISPLAY_NAME = 'Inter-­European Aid Association';

var SUBJECT_DE_SEATS = 'Bestehendes Team - Sitzplatz verfügbar';
var SUBJECT_EN_SEATS = 'Existing Team - seat available';
var SUBJECT_DE_NO_SEATS = 'Bestehendes Team - kein Sitzplatz verfügbar';
var SUBJECT_EN_NO_SEATS = 'Existing Team - no seat available';

var BODY_DE_SEATS = '1l0a0bxGrIPTjN4-F5CFFaF-9GDT_PuP_7haE1HyTscE';
var BODY_EN_SEATS = '1tEFqUGBrfxZ550vLghm---cqgCb_ikgXRpTXWX1autQ';
var BODY_DE_NO_SEATS = '1tduMNcMu7r0N1bbih5x5eljwmEB9RCNCpKHFAFzj9UQ';
var BODY_EN_NO_SEATS = '1X8CGJcZ7hY7J4_O9jHVV08zTfc3pdN_G9j1aDPjQmFE';

var ATTACHMENTS_PACKAGING = '0By52Ty8g_M6JVDBCc24wUk5Ra1k';
var ATTACHMENTS_VOLUNTEER_INFO = '0By52Ty8g_M6JYXRsMUZLUGdRVm8';
var ATTACHMENTS_LABELS = '0By52Ty8g_M6JUHZQdW4wclhqVGs';

var TEAMS_SEATS = 'Free seats';

var LANGUAGE = 'Preferred communication language';
var LANGUAGE_ENGLISH = 'English';
var LANGUAGE_GERMAN = 'German';

var FORM_EMAIL = 9.57287598E8;
var FORM_LANGUAGE = 2.047945738E9;
var FORM_SEATS = 2.90550958E8;

function test() {
  var range = SpreadsheetApp.openById('1j2xLQ8Z1sa0ny2mp77Ipu1-LEDgyuubUWFhZalQ9Dqk').getSheetByName(SHEET_TEAMS).getRange('a1');
  var values = {
    'E-Mail address': ['kugelmann.dennis@gmail.com'],
    'Transportation': [2],
    'Free seats': ['0'],
    'Preferred communication language': [LANGUAGE_ENGLISH]
  }
  var e = {
    'range': range,
    'namedValues': values
  }
  onFormSubmit(e);
}

function onFormSubmit(e) {
  var form = FormApp.getActiveForm();
  var emailItem = form.getItemById(FORM_EMAIL);
  var languageItem = form.getItemById(FORM_LANGUAGE);
  var seatsItem = form.getItemById(FORM_SEATS);
  
  var emailResponse = e.response.getResponseForItem(emailItem);
  var languageResponse = e.response.getResponseForItem(languageItem);
  var seatsResponse = e.response.getResponseForItem(seatsItem);
  
  var email = emailResponse.getResponse();
  var language = languageResponse.getResponse();
  var seats = seatsResponse.getResponse();
  
  var packaging = DriveApp.getFileById(ATTACHMENTS_PACKAGING);
  var volunteerInfo = DriveApp.getFileById(ATTACHMENTS_VOLUNTEER_INFO);
  var labels = DriveApp.getFileById(ATTACHMENTS_LABELS);
  
  var attachments = [packaging.getAs(MimeType.PDF), 
                     volunteerInfo.getAs(MimeType.PDF), 
                     labels.getAs(MimeType.PDF)];
  //Send team files
  if (seats > 0) {
    //Send free seats
    if (language == LANGUAGE_GERMAN) {
      //Send german version
      sendEmail(email, SUBJECT_DE_SEATS, BODY_DE_SEATS, attachments);
      Logger.log('team > free seats > german');
    } else {
      //Send english version
      sendEmail(email, SUBJECT_EN_SEATS, BODY_EN_SEATS, attachments);
      Logger.log('team > free seats > english');
    }
  } else {
    //Send no seats
    if (language == LANGUAGE_GERMAN) {
      //Send german version
      sendEmail(email, SUBJECT_DE_NO_SEATS, BODY_DE_NO_SEATS, attachments);
      Logger.log('team > no free seats > german');
    } else {
      //Send english version
      sendEmail(email, SUBJECT_EN_NO_SEATS, BODY_EN_NO_SEATS, attachments);
      Logger.log('team > no free seats > english');
    }
  }
}

function sendEmail(email, subject, contentID, attachments) {
  var content = DocumentApp.openById(contentID)
                                 .getBody()
                                 .editAsText()
                                 .getText();
  MailApp.sendEmail(email, subject, content, {
    attachments: attachments,
    name: EMAIL_DISPLAY_NAME,
    replyTo: EMAIL_REPLY_TO_ADDRESS
  });
  
  var ss = SpreadsheetApp.openById(SPREADSHEET);
  var sheet = ss.getSheetByName(SHEET_TEAMS);
  var emailColumn = findColumn(sheet, EMAIL_ADDRESS_FIELD);
  var row = findEmailRow(sheet, emailColumn, email);
  var column = findColumn(sheet, INFO_FLYER);
  sheet.getRange(row, column).setValue(INFO_FLYER_YES);
}

function findColumn(sheet, type) {
  var range = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  for(i=0; i<range[0].length; i++) {
    if (range[0][i] === type) {
      return i+1;
    }
  }
  return -1;
}

function findEmailRow(sheet, column, email) {
  var range = sheet.getRange(2, column, sheet.getLastRow()-1).getValues();
  for(n=range.length-1; n>=0; n--) {
    if(range[n][0] === email) {
      return n+2
    }
  }
}

function log() {
  var form = FormApp.getActiveForm();
  var items = form.getItems();
  for(i=0; i<items.length; i++) {
    var title = items[i].getTitle();
    var id = items[i].getId();
    Logger.log(title);
    Logger.log(id);
  }
}
