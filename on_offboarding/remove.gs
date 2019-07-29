function remove() {
  try{
    
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Offboarding");
    var cell_range = activeSheet.getRange("F9:H11");
    var error_1 = activeSheet.getRange(9, 8);
    var error_2 = activeSheet.getRange(10, 8);
    var error_3 = activeSheet.getRange(11, 8);
    
    var ui = SpreadsheetApp.getUi();
    var username = ui.prompt('Please Provide Your Atlassian Account Credentials', 'Username', ui.ButtonSet.YES_NO);
    
    // Process the user's response.
    if (username.getSelectedButton() == ui.Button.YES) {
      
      Logger.log('The user\'s name is %s.', username.getResponseText());
      var password = ui.prompt('Please Provide Your Atlassian Account Credentials', 'Password', ui.ButtonSet.YES_NO);
      
      if (password.getSelectedButton() == ui.Button.YES) {
        
      } else {
        Logger.log('The user clicked the close button in the dialog\'s title bar.');
      }
      
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
    }
    
    var userEmail = Session.getActiveUser().getEmail();
    Logger.log("Executed by: "+userEmail);
    
    var table = activeSheet.getRange(8,2,16,2).getValues();
    var primaryEmailAddress = activeSheet.getRange(4,2).getValue();
    
    cell_range.clear();
    
    cell_range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    cell_range.setFontFamily("Cambria");
    cell_range.setFontSize(12);
    
    var currentUserStatus =  activeSheet.getRange(9, 6);
    var removeUserfromGroupStatus =  activeSheet.getRange(10, 6);
    var emailStatusOff =  activeSheet.getRange(11, 6);
    
    
    currentUserStatus.setValue("User found").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    removeUserfromGroupStatus.setValue("Remove from groups").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    emailStatusOff.setValue("Email Sent").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    
    
    
    setStatus_r("Pending", "userStatus");
    setStatus_r("Pending", "removeuserfromGroup");
    setStatus_r("Pending", "emailSentOff");
    
    HR_TEAM = activeSheet.getRange(4,6).getValue();
    HR_ID = activeSheet.getRange(5,6).getValue();
    
    var userdetails = removefromgroup(primaryEmailAddress, table[1][1],table[2][1],table[4][1], table[5][1], table[6][1], table[8][1], table[9][1], table[14][1], table[15][1], HR_TEAM, HR_ID, table[10][1] , username , password);
    
  } catch (e) {
    Logger.log(e);
  }
}

var currentUserStatus =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Offboarding").getRange(9, 7);
var removeUserfromGroupStatus =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Offboarding").getRange(10, 7);
var emailStatusOff =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Offboarding").getRange(11, 7);


function setStatus_rColor(status, cell){
  switch(status){
    case "Pending":
      cell.setBackground("#ff9900");
      break;
    
    case "Success":
      cell.setBackground("green");
      break;
      
    case "Failed":
      cell.setBackground("red");
      break;
    
  }
}

function setStatus_r(status, action){
  
  switch(action){
      
    case "userStatus":
      currentUserStatus.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      setStatus_rColor(status, currentUserStatus)
      currentUserStatus.setFontColor('white');
      break;
      
    case "removeuserfromGroup":
      removeUserfromGroupStatus.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      setStatus_rColor(status, removeUserfromGroupStatus)
      removeUserfromGroupStatus.setFontColor('white');
      break;
      
    case "emailSentOff":
      emailStatusOff.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      setStatus_rColor(status, emailStatusOff)
      emailStatusOff.setFontColor('white');
      break;
      
  }
}

function removefromgroup(primaryEmailAddress, firstName, lastName, secondaryEmailAddress, functions, department, location, country, org, state, hrTeam, hrId, et , username , password) {
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Offboarding");
  var cell_range = activeSheet.getRange("F9:H11");
  var error_1 = activeSheet.getRange(9, 8);
  var error_2 = activeSheet.getRange(10, 8);
  var error_3 = activeSheet.getRange(11, 8);
  
  var user_name = primaryEmailAddress.split("@", 2)[0];
  
  try{
    var url = "https://reancloud.atlassian.net/rest/api/2/user?username="+user_name;
    
    var headers = {
      "contentType": "application/json",
      "headers":{
        "User-Agent": "Group Finder",
        "Authorization": "Basic " + Utilities.base64Encode(username.getResponseText() + ":" + password.getResponseText())
      },
      "validateHttpsCertificates" :false
    };
    
    var response = UrlFetchApp.fetch(url, headers);
    var fact = response.getContentText();
    var dataAll = JSON.parse(fact);
    
    if( dataAll.name ){
      setStatus_r("Success", "userStatus");
    }
    
    try{
      
      removeGroupMember(primaryEmailAddress , username , password)
      
      setStatus_r("Success", "removeuserfromGroup");
    } catch (e) {
      Logger.log(e)
      setStatus_r("Failed", "removeuserfromGroup");
      error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);    
      
    };
    
    return dataAll.name
   
    
  } catch (e) {
    setStatus_r("Failed", "userStatus");
    error_1.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);    
    
    
    Logger.log(e)
  };
  
}

function removeGroupMember(userEmail , username , password) {

  var user_name = userEmail.split("@", 2)[0];
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Offboarding");
  var cell_range = activeSheet.getRange("F9:H11");
  var error_1 = activeSheet.getRange(9, 8);
  var error_2 = activeSheet.getRange(10, 8);
  var error_3 = activeSheet.getRange(11, 8);
  var b = [];
  
  try{ 
    var url = "https://reancloud.atlassian.net/rest/api/2/user?username="+user_name+"&expand=groups";
    
    var headers = {
      "contentType": "application/json",
      "headers":{
        "User-Agent": "Group Finder",
        "Authorization": "Basic " + Utilities.base64Encode(username.getResponseText() + ":" + password.getResponseText())
      },
      "validateHttpsCertificates" :false
    };
    
    var response = UrlFetchApp.fetch(url, headers);
    var fact = response.getContentText();
    var dataAll = JSON.parse(fact);
    
    var users_data = dataAll.groups.items;
    for (i in users_data) {
      var group_name = users_data[i].name;
      b.push(group_name);
      var formData = {
        'username': user_name,
        'groupname' : group_name
      };
      
      var header = {
        "Authorization": "Basic " + Utilities.base64Encode(username.getResponseText() + ":" + password.getResponseText())
      };
      var options = {
        "method" : 'delete',
        "payload" : JSON.stringify(formData),
        'contentType': 'application/json',   
        "headers": header
      };
      var url = "https://reancloud.atlassian.net/rest/api/2/group/user?username="+user_name+"&groupname="+group_name;
      var response = UrlFetchApp.fetch(url, options); 
    }
    
    setStatus_r("Success", "removeuserfromGroup"); 

        emailNotification_r(userEmail , b , username);
    
  } catch (e) {
    Logger.log(e)
    setStatus_r("Failed", "removeuserfromGroup");
    error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);    
    
  }
  
}


function emailNotification_r(userID , grouList , username ) {
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Offboarding");
  var cell_range = activeSheet.getRange("F9:H11");
  var error_1 = activeSheet.getRange(9, 8);
  var error_2 = activeSheet.getRange(10, 8);
  var error_3 = activeSheet.getRange(11, 8);  
  
  try{
    
    var email_cc = "rean.onboarding@hitachivantara.com, rean_globalhrops@hitachivantara.com"
//    var email_cc = "nikhil.kapure@reancloud.com"
    
    var email_to = Session.getActiveUser().getEmail();
    var date = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy");
    var group_count = grouList.length;
    var user_count = userID.length;
    var executerID = Session.getActiveUser().getEmail();
    var atlassianID = username.getResponseText();
    var massage = "";

    var subject = "[Atlassian Offboarding] User "+userID+ " removed from "+group_count+ " Atlassian groups on "+date ;     
    
    if (group_count == "0"){
      var massage = "User "+userID+" not belong from any groups." 
    } else {
    var massage = "Today [ "+date+" ] "+userID+ " has been offboarded from Atlassian sucessfully. That user is removed from the following groups"
    }
    
    var htmlTemplate_add = HtmlService
    .createTemplateFromFile("onandoffboarding");
    htmlTemplate_add.data = [userID , grouList, date , executerID , atlassianID , massage];
    
    var htmlBody = htmlTemplate_add.evaluate().getContent();
    MailApp.sendEmail({
      to: email_to,
      cc: email_cc,
      subject: subject,
      htmlBody: htmlBody
    });   
    
    setStatus_r("Success", "emailSentOff");
    sendAcknowledgement();
    
  } catch (e) {
    Logger.log(e.toString()); 
    setStatus_r("Failed", "emailSentOff");
    error_3.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    return false;
    
  }
  
}

function sendAcknowledgement(){
  try{
    
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Offboarding");
    
    var table = activeSheet.getRange(8,2,16,2).getDisplayValues();
    
    var employeeName = table[3][1];
    var functions = table[5][1];
    var department  = table[6][1];
    var role = table[7][1];
    var location = table[8][1];
    var country = table[9][1];
    var engagementType = table[10][1];
    
    var htmlTemplate = HtmlService
    .createTemplateFromFile("acknowledgement");
    htmlTemplate.data = [employeeName, location,department,functions];
    var htmlBody = htmlTemplate.evaluate().getContent();
    
    var subject = "[Offboarding acknowledgement] The Offboarding Process has been initiated for employee "+employeeName;  
    var email_to = "rean.offboarding@hitachivantara.com,rean.jira.admin@hitachivantara.com";
    
    MailApp.sendEmail({
      to: email_to,
      cc: "rean.global.hr.operations@hitachivantara.com",
      subject: subject,
      htmlBody: htmlBody
    });
    
  } catch (e) {
    Logger.log(e)
  }
}
