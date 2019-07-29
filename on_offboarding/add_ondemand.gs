function master() {
  try{
    
    var userEmail = Session.getActiveUser().getEmail();
    Logger.log("Exeted by: "+userEmail);
   
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
    
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("On-Demand");
    
    var group_list = activeSheet.getRange("B2:B").getValues().filter(String);
    var user_id = activeSheet.getRange("A2:A").getValues().filter(String);
    var cell_range = activeSheet.getRange("E3:G5");
    var error_1 = activeSheet.getRange(3, 7);
    var error_2 = activeSheet.getRange(4, 7);
    var error_3 = activeSheet.getRange(5, 7);

    cell_range.clear();
    
    cell_range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    cell_range.setFontFamily("Cambria");
    cell_range.setFontSize(12);
    
    var cU =  activeSheet.getRange(3, 5);
    var aU =  activeSheet.getRange(4, 5);
    var eU =  activeSheet.getRange(5, 5);
    
    cU.setValue("User Account Create").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    aU.setValue("Add User in Groups").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    eU.setValue("Email Sent").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    
    changeStatus("Pending", "createUserJia");
    changeStatus("Pending", "addUserinJiraGroup");    
    changeStatus("Pending", "emailSent");    
    
    var a_1 = jiraAdd(group_list , user_id, activeSheet , error_1 , error_2 , error_3 , username , password);
  } catch (e) {
    Logger.log(e);
  }
}



function changeStatusColor(status, cell){
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

function changeStatus(status, action){
  
  var cU =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("On-Demand").getRange(3, 6);
  var aU =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("On-Demand").getRange(4, 6);
  var eU =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("On-Demand").getRange(5, 6);
  
  switch(action){
      
    case "createUserJia":
      cU.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      changeStatusColor(status, cU)
      cU.setFontColor('white');
      break;
      
    case "addUserinJiraGroup":
      aU.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      changeStatusColor(status, aU)
      aU.setFontColor('white');
      break;
      
    case "emailSent":
      eU.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      changeStatusColor(status, eU)
      eU.setFontColor('white');
      break;
      
  }
}

function addinGroups(user_name , group_list , activeSheet , error_1 , error_2 , error_3 , a , username , password) {

  var b = [];
  
  try{
    for(var group in group_list){
      var groupEmail = group_list[group];
      var group_name = groupEmail[0];
      b.push(group_name)
      
      try{
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
        
        var url = "https://reancloud.atlassian.net/rest/api/2/group/user?groupname="+group_name;
        var response = UrlFetchApp.fetch(url, options);     
        
      } catch (e) {
        
        Logger.log(e)
        error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
        throw e;
      }
    }
    
  } catch (e) {
    Logger.log(e)
    changeStatus("Failed", "addUserinJiraGroup");
    error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    throw e;
  }     
  
  return true;  
  
}   

function jiraAdd(group_list , user_id , activeSheet , error_1 , error_2 , error_3 , username , password) {
 
  var a = [];
  
  var userJiaStatus = false;
  try{
    
    changeStatus("Pending", "createUerJia");
    changeStatus("Pending", "addUserinJiraGroup");
    changeStatus("Pending", "emailSent");
    
    //Check user status
    try{
      for(var i in user_id){
        var userId = user_id[i];
        a.push(userId); 
        var user_name = userId[0].split("@")[0];
        
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
        
      }
      changeStatus("Success", "createUserJia");
      userJiaStatus = true;
      Logger.log("createUserJia:success");
    }catch (e) {
      Logger.log(e)
      changeStatus("Failed", "createUserJia");
      error_1.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
      return false;
    } 
    
    Logger.log("createUserJia:success:process addUserinJiraGroup");
    var addInGroupStatus = false;
    if( userJiaStatus ){
      try{
        for(var i in user_id){
          var userId = user_id[i];
          var user_name = userId[0].split("@")[0];
          addinGroups(user_name , group_list , activeSheet , error_1 , error_2 , error_3 , a , username , password);
        }
        
        changeStatus("Success", "addUserinJiraGroup");
        addInGroupStatus = true;
        Logger.log("addUserinJiraGroup:success");
        
      }catch (e) {
        Logger.log(e)
        changeStatus("Failed", "addUserinJiraGroup");
        error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
        return false;
      }
      
    }
    
    Logger.log("addUserinJiraGroup:success:process emailNotification");
    if ( addInGroupStatus){
      emailSentAdd(a , group_list , activeSheet , error_1 , error_2 , error_3 , username);  
    }
    
  } catch (e) {
    Logger.log(e)
    error_1.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    return false;
  }   
}
  
function  emailSentAdd(a, b , activeSheet , error_1 , error_2 , error_3 , username) {
  
  try{
    
    var email_cc = "GlobalHROps@reancloud.com"
//    var email_cc = "nikhil.kapure@reancloud.com"
    
    var email_to = Session.getActiveUser().getEmail();
    var date = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy");
    var user_count = a.length;
    var group_count = b.length;
    var atlassianID = username.getResponseText();
    
    var subject = "[Atlassian Group] Total "+user_count+ " members Added in "+group_count+ " Atlassian groups on "+date ;     
    
    var massage = "Total "+user_count+ " members Added in "+group_count+ " Atlassian groups on "+date+". Please refer following table for more details."; 
    
    
    var htmlTemplate_add = HtmlService
    .createTemplateFromFile("ondemand");
    htmlTemplate_add.data = [a , b, user_count > group_count ? user_count : group_count, email_to , atlassianID , massage];
    
    var htmlBody = htmlTemplate_add.evaluate().getContent();
    MailApp.sendEmail({
      to: email_to,
      cc: email_cc,
      subject: subject,
      htmlBody: htmlBody
    }); 
    changeStatus("Success", "emailSent");
    
  } catch (e) {
    Logger.log(e.toString()); 
    changeStatus("Failed", "emailSent");
    error_3.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);    
    return false;
    
  }
  
}
