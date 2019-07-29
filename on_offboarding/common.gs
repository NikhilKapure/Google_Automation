var WELCOME_MAIL = "welcomeMail";
var CREATE_USER = "createUser";
var PASSWORD_SENT = "passwordSent";
var ADD_USER_TO_GROUP = "adduserToGroup";

var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Onboarding");
var cell_range = activeSheet.getRange("F9:H11");
var error_1 = activeSheet.getRange(9, 8);
var error_2 = activeSheet.getRange(10, 8);
var error_3 = activeSheet.getRange(11, 8);

var createUserStatus =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Onboarding").getRange(9, 7);
var addUserToGroupStatus =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Onboarding").getRange(10, 7);
var emailsentStatus =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Onboarding").getRange(11, 7);

function selectUserGroupDLs(group, user, country, engType, state){
  
  if( country === "US" || country === "Australia" || country === "Israel" || country === "UK" || country === "Poland" || country === "Kenya") {
    
    var account_management = ["confluence-users" , "jira-users" , "global-internal-engineering", "confluence-AM-admin"];
    var delivery_appdev = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var delivery_bigdata = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var delivery_devops = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var managed_services = ["confluence-users" , "jira-users" , "global-internal-MGS" , "global-internal-engineering" , "global-internal-architects", "confluence-MAN-rw"];
    var pmo = ["confluence-users" , "jira-users" , "global-internal-pmo" , "global-internal-engineering"];
    var product_development = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var professional_services = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var finance = ["confluence-users" , "jira-users" , "Global-internal-finance-US" , "confluence-FIN-ro" , "confluence-FIN-rw" , "confluence-FIN-admin" , "global-internal-finance-47Lining"];
    var hr_grp = ["confluence-users" , "jira-users" , "global-internal-hr" , "confluence-REC-rw"];
    var operations = ["confluence-users" , "jira-users" , "global-internal-hr" , "confluence-REC-rw"];
    var marketing = ["confluence-users" , "jira-users" , "global-internal-marketing" , "confluence-MAR-ro" , "confluence-MAR-rw" , "confluence-MAR-admin"];
    var sales = ["confluence-users" , "jira-users" , "global-internal-sales" , "confluence-SALES-ro" , "confluence-SALES-rw" , "confluence-SALES-admin"];
    
    var employees = ["global-internal-employees"];
    var interns = ["global-internal-interns"];
    var consultants = ["global-internal-consultants"];  
    
    Logger.log("Mapping goups as per the US standerd")
  }
  
  else if( country === "India" ) {     
    
    var account_management = ["confluence-users" , "jira-users" , "global-internal-engineering", "confluence-AM-admin"];
    var delivery_appdev = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var delivery_bigdata = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var delivery_devops = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var managed_services = ["confluence-users" , "jira-users" , "global-internal-MGS" , "global-internal-engineering" , "global-internal-architects", "confluence-MAN-rw"];
    var pmo = ["confluence-users" , "jira-users" , "global-internal-pmo" , "global-internal-engineering"];
    var product_development = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var professional_services = ["confluence-users" , "jira-users" , "global-internal-engineering" , "global-internal-architects"];
    var finance = ["confluence-users" , "jira-users" , "global-internal-finance-India" , "confluence-FIN-ro" , "confluence-FIN-rw" , "confluence-FIN-admin"];
    var hr_grp = ["confluence-users" , "jira-users" , "global-internal-hr" , "confluence-REC-rw"];
    var operations = ["confluence-users" , "jira-users" , "global-internal-hr" , "confluence-REC-rw"];
    var marketing = ["confluence-users" , "jira-users" , "global-internal-marketing" , "confluence-MAR-ro" , "confluence-MAR-rw" , "confluence-MAR-admin"];
    var sales = ["confluence-users" , "jira-users" , "global-internal-sales" , "confluence-SALES-ro" , "confluence-SALES-rw" , "confluence-SALES-admin"];
    
    var employees = ["global-internal-employees"];
    var interns = ["global-internal-interns"];
    var consultants = ["global-internal-consultants"];      
    
    Logger.log("Mapping goups as per the INDIA standerd")
  }
  
  else{
    
    Logger.log("No India No US")
    
  }      
  
  var groups;
  switch(group) {
    case "Account_Management":
      Logger.log("Adding to account_management");
      groups = account_management;
      break;
      
    case "Delivery (AppDev)":
      Logger.log("Adding to delivery_appdev");
      groups = delivery_appdev;
      break;
      
    case "Delivery (BigData)":
      Logger.log("Adding to delivery_bigdata");
      groups = delivery_bigdata; 
      break;
      
    case "Delivery (DevOps)":
      Logger.log("Adding to delivery_devops");
      groups = delivery_devops;
      break;
      
    case "Managed_Services":
      Logger.log("Adding to managed_services");
      groups = managed_services;
      break;
      
    case "PMO":
      Logger.log("Adding to pmo");
      groups = pmo;
      break;
      
    case "Product_Development":
      Logger.log("Adding to product_development");
      groups = product_development;
      break;
      
    case "Professional_Services":
      Logger.log("Adding to professional_services");
      groups = professional_services;
      break;
      
    case "Finance":
      Logger.log("Adding to finance");
      groups = finance;
      break;
      
    case "HR":
      Logger.log("Adding to hr");
      groups = hr_grp;
      break;
      
    case "Operations":
      Logger.log("Adding to operations");
      groups = operations;
      break;
      
    case "Marketing":
      Logger.log("Adding to marketing");
      groups = marketing;
      break;
    case "Sales":
      Logger.log("Adding to sales");
      groups = sales;
      break;
      
    case "Employee":
      Logger.log("Adding to employees");
      groups = employees;
      break;
      
    case "Intern":
      Logger.log("Adding to interns");
      groups = interns;
      break;
      
    case "Consultant":
      Logger.log("Adding to consultants");
      groups = consultants;
      break
      
      default:
      Logger.log("Target DL Group not found")
  }
  
  return groups;
}


function setStatusColor(status, cell){
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

function setStatus(status, action){
  
  switch(action){
      
    case "createUser":
      createUserStatus.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      setStatusColor(status, createUserStatus)
      createUserStatus.setFontColor('white');
      break;
      
    case "adduserToGroup":
      addUserToGroupStatus.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      setStatusColor(status, addUserToGroupStatus)
      addUserToGroupStatus.setFontColor('white');
      break;
      
    case "emailsentonStatus":
      emailsentStatus.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      setStatusColor(status, emailsentStatus)
      emailsentStatus.setFontColor('white');
      break;        
      
  }
}

function addInJira(primaryEmailAddress, firstName, lastName, secondaryEmailAddress, functions, department, location, country, org, state, hrTeam, hrId, et  , username , password) {
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Onboarding");
  var cell_range = activeSheet.getRange("F9:H11");
  var error_1 = activeSheet.getRange(9, 8);
  var error_2 = activeSheet.getRange(10, 8);
  var error_3 = activeSheet.getRange(11, 8);
  
  //  var scriptProperties = PropertiesService.getScriptProperties();
  //  var data = scriptProperties.getProperties();
  var user_name = primaryEmailAddress.split("@", 2)[0];
  var userJiaStatus = false;
  
  try{
    var url = "https://reancloud.atlassian.net/rest/api/2/user?username="+user_name;
    var userStatus = "Failed";
    
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
      setStatus("Success", "createUser");  
      userJiaStatus = true;
      Logger.log("User ID is present : success");
      removeFromRespectiveGroups(user_name, username, password);
      
    } 
    
    var groupsToAddUser = [];
    
    var addInGroupStatus = false;
    if( userJiaStatus ){
      
      try{
        
        var groups = selectUserGroupDLs(department, primaryEmailAddress, country, et, state);
        for(var i in groups){
          groupsToAddUser.push(groups[i]);
        }
        
        groups = selectUserGroupDLs(et, primaryEmailAddress, country, et, state);
        for(var i in groups){
          groupsToAddUser.push(groups[i]);
          
        }
        
        addGroupMember(groupsToAddUser, primaryEmailAddress , username , password);
        setStatus("Success", "adduserToGroup");
        addInGroupStatus = true;
        Logger.log("Added in to the groups : success");
        
        
      }catch (e) {
        Logger.log(e)
        setStatus("Failed", "adduserToGroup");
        error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
        return false;
      }
      
    }
    
    Logger.log("adduserToGroup : success : process emailSentonboarding");
    if ( addInGroupStatus){
      emailSentonboarding(primaryEmailAddress , groupsToAddUser , username);    
    }    
    
    
  } catch (e) {
    Logger.log(e)
    setStatus("Failed", "createUser");
    error_1.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
  };
  
}
  

function addGroupMember(groups, userEmail , username , password) {
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Onboarding");
  var cell_range = activeSheet.getRange("F9:H11");
  var error_1 = activeSheet.getRange(9, 8);
  var error_2 = activeSheet.getRange(10, 8);
  var error_3 = activeSheet.getRange(11, 8);
  
  var user_name = userEmail.split("@", 2)[0];
  
  
  var formData = {
    'name': user_name
  };
  
  var header = {
    "Authorization": "Basic " + Utilities.base64Encode(username.getResponseText() + ":" + password.getResponseText())
  };
  var options = {
    "method" : 'post',
    "payload" : JSON.stringify(formData),
    'contentType': 'application/json',   
    "headers": header
  };
  
  
  try{
    
    for(group in groups){
      
      var groupEmail = groups[group];
      var url = "https://reancloud.atlassian.net/rest/api/2/group/user?groupname="+groupEmail;
      var response = UrlFetchApp.fetch(url, options); 
      
    }
    
    setStatus("Success", "adduserToGroup"); 
    
  } catch (e) {
    Logger.log(e)
    setStatus("Failed", "adduserToGroup");
    error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    throw e;
  }     
  
  return true;  
}

function  emailSentonboarding(userID , grouList , username) {
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Onboarding");
  var cell_range = activeSheet.getRange("F9:H11");
  var error_1 = activeSheet.getRange(9, 8);
  var error_2 = activeSheet.getRange(10, 8);
  var error_3 = activeSheet.getRange(11, 8);  
  
  
  try{
    
    var email_cc = "rean.onboarding@hitachivantara.com,rean_globalhrops@hitachivantara.com"
//    var email_cc = "nikhil.kapure@reancloud.com"
    
    var email_to = Session.getActiveUser().getEmail();
    var date = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy");
    var group_count = grouList.length;
    var user_count = userID.length;
    var executerID = Session.getActiveUser().getEmail();
    var atlassianID = username.getResponseText();
    
    var subject = "[Atlassian Onboarding] User "+userID+ " onboarded sucessfully in Atlassian on "+date ;     
    
    var massage = "Today [ "+date+" ] "+userID+ " has been onboarded in Atlassian sucessfully. That user is added in the following groups"
    
    
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
    setStatus("Success", "emailsentonStatus");
    
  } catch (e) {
    Logger.log(e.toString()); 
    setStatus("Failed", "emailsentonStatus");
    error_3.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);    
    return false;
    
  }
  
}

function removeFromRespectiveGroups(user_name , username , password) {
  Logger.log("Deleteing user from grups")
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
    
  }catch(e){
    Logger.log(e)
  } 
}
