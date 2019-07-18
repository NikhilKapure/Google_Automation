/* 


  ======================================================
  
     Laptop Allocation & Asset Tracker Sheet - India

  ======================================================
  
  Written by Nikhil Kapure on 06/07/2018 
  
*/

function clickhere() {
// 1 = row 1, 14 = column 14 = N
goToSheet("Asset_List", 1, 1);
}

function goToSheet(sheetName, row, col) {
var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
SpreadsheetApp.setActiveSheet(sheet);
//var range = sheet.getRange(row, col)
var range = sheet.getRange(sheet.getLastRow(), 1);
SpreadsheetApp.setActiveRange(range);
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('REANCloud IT Operations')
      .addItem('Asset Allocation', 'allocation')  
      .addItem('Asset Submission', 'submission')
      .addToUi();
} 


function allocation() {
  
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to send email Notification?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes"
  
  var userEmail = Session.getActiveUser().getEmail(); 
  Logger.log("Executed by: "+userEmail);

  var activeSheet = SpreadsheetApp.getActive().getSheetByName("Main Menu");
  
  var hr = "rean.india.hr@hitachivantara.com";
//  var hr = "nikhil.kapure@reancloud.com";
  var hyd_support = "sivasankar.nagireddy@reancloud.com";
//  var hyd_support = "nikhil.kapure@reancloud.com";

  var global_support = "nikhil.kapure@reancloud.com";
  var userName = activeSheet.getRange("F3").getDisplayValue();
  var serviceTag = activeSheet.getRange("F13").getDisplayValue();
  var userEmailID = activeSheet.getRange("I14").getDisplayValue();
  var manufacturer = activeSheet.getRange("F7").getDisplayValue();
  var model = activeSheet.getRange("F8").getDisplayValue();
  var jiraTicket = activeSheet.getRange("I15").getDisplayValue();
  var menu = activeSheet.getRange("E2:E21").getDisplayValues().filter(String);
  var info = activeSheet.getRange("F2:F21").getDisplayValues().filter(String);
  var comments = activeSheet.getRange("I16").getDisplayValue();
  var date = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy");
  var status = "Allocation"
  if (userEmail === hyd_support ){
    
    Logger.log("not Nikhil")
    
    var email_to = userEmailID    
    var email_cc = global_support +","+ hr;    
    
  } else {
    
    Logger.log("its Nikhil")
    
    var email_to = userEmailID    
    var email_cc = hyd_support +","+ hr;
  }
  
  var msg = manufacturer+" "+model+" "+serviceTag+" asset has been assigned to "+userName+ " on "+ date;

  var subject = "[Assets Management] "+msg; 

  sendEmail(msg, jiraTicket,userName,menu,info,email_to,email_cc,subject, comments, status);
  onOpen()
  
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Email Notification Sending has been cancelled.');
  }
}

function submission() {

    var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to send email Notification?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes"
    
    var userEmail = Session.getActiveUser().getEmail(); 
  Logger.log("Executed by: "+userEmail);

  var activeSheet = SpreadsheetApp.getActive().getSheetByName("Main Menu");
  
  var hr = "rean.india.hr@hitachivantara.com";
//  var hr = "nikhil.kapure@reancloud.com";
  var hyd_support = "sivasankar.nagireddy@reancloud.com";
//  var hyd_support = "nikhil.kapure@reancloud.com";

  var global_support = "nikhil.kapure@reancloud.com";
  var userName = activeSheet.getRange("F3").getDisplayValue();
  var serviceTag = activeSheet.getRange("F13").getDisplayValue();
  var userEmailID = activeSheet.getRange("I14").getDisplayValue();
  var manufacturer = activeSheet.getRange("F7").getDisplayValue();
  var model = activeSheet.getRange("F8").getDisplayValue();
  var jiraTicket = activeSheet.getRange("I15").getDisplayValue();
  var comments = activeSheet.getRange("I16").getDisplayValue();
  var date = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy");
  var status = "Submission"
  var menu = activeSheet.getRange("E2:E21").getDisplayValues().filter(String);
  var info = activeSheet.getRange("F2:F21").getDisplayValues().filter(String);
  
  if (userEmail === hyd_support ){
    
    Logger.log("not Nikhil")
    
    var email_to = userEmailID    
    var email_cc = global_support +","+ hr;    
    
  } else {
    
    Logger.log("its Nikhil")
    
    var email_to = userEmailID    
    var email_cc = hyd_support +","+ hr;
  }
  
  var msg = "This is to acknowledge the receipt of below asset from you and here are the details -";  
 
  var subject = "[IT Clearance] Submission of Assets Statement - "+userName+" on " +date; 

  sendEmail(msg, jiraTicket,userName,menu,info,email_to,email_cc,subject, comments, status);
  onOpen();
  
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Email Notification Sending has been cancelled.');
  }  
}

function sendEmail(msg, jiraTicket,userName,menu,info,email_to,email_cc,subject, comments, status){
 
  Logger.log(email_to)
  Logger.log(email_cc)

  try{
    var htmlTemplate = HtmlService
    .createTemplateFromFile("allocation");
    htmlTemplate.data = [msg, jiraTicket,userName,menu,info,comments];
    var htmlBody = htmlTemplate.evaluate().getContent();
    MailApp.sendEmail({
      to: email_to,
      cc: email_cc,
      subject: subject,
      htmlBody: htmlBody
    });
    
  var ui = SpreadsheetApp.getUi(); // Same variations.
    ui.alert('Thank you! The email notification has been sent successfully.');

    
  } catch (e) {
    Logger.log(e)
  }

 uploadPDF(msg, jiraTicket,userName,menu,info,email_to,email_cc,subject, comments, status);
}

function uploadPDF(msg, jiraTicket,userName,menu,info,email_to,email_cc,subject, comments, status){

   try{
     var htmlTemplate = HtmlService
     .createTemplateFromFile("allocation");
     htmlTemplate.data = [msg, jiraTicket,userName,menu,info,comments];
     var htmlBody = htmlTemplate.evaluate().getContent();
     
     var blob = Utilities.newBlob(htmlBody, MimeType.HTML);
     blob = blob.getAs(MimeType.PDF).setName(subject+".pdf");
     Logger.log("File converted in PDF sucessfully")
     
     if (status === "Allocation"){
       var destination_id = "1LtW0XM4ika9se0PFLu0ocEvR0KSRrIOH";
       
     } else if (status === "Submission") {
       
       var destination_id = "1M99bLE0rQdHNFK4jc6p1qsABzElORLxh";
       
     }
     
  var contentType = "application/pdf";
  var destination = DriveApp.getFolderById(destination_id);
  destination.createFile(blob); 
  Logger.log("File Uploaded Sucessfully")
  var ui = SpreadsheetApp.getUi(); // Same variations.
    ui.alert('File has been successfully uploaded on Google Drive.');    
  } catch (e) {
    Logger.log(e)
  }
 
}
