function slave() {
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
    var cell_range = activeSheet.getRange("E3:G5");
    var error_1 = activeSheet.getRange(3, 7);
    var error_2 = activeSheet.getRange(4, 7);
    var error_3 = activeSheet.getRange(5, 7);
    
    var group_list = activeSheet.getRange("B2:B").getValues().filter(String);
    var user_id = activeSheet.getRange("A2:A").getValues().filter(String);
    
    cell_range.clear();
    
    cell_range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    cell_range.setFontFamily("Cambria");
    cell_range.setFontSize(12);
    
    var cU =  activeSheet.getRange(3, 5);
    var aU =  activeSheet.getRange(4, 5);
    var eU =  activeSheet.getRange(5, 5);
    
    cU.setValue("User found").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    aU.setValue("Remove from groups").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    eU.setValue("Email Sent").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    
    changeStatus_r("Pending", "userJia");
    changeStatus_r("Pending", "removeFromGroup");    
    changeStatus_r("Pending", "emailSent");    
    
    
    var v = jiraRemove(group_list , user_id , activeSheet , error_1 , error_2 , error_3 , username , password);
  } catch (e) {
    Logger.log(e);
  }
}



function changeStatus_rColor(status, cell){
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

function changeStatus_r(status, action){
  
  var cUs =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("On-Demand").getRange(3, 6);
  var aUs =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("On-Demand").getRange(4, 6);
  var eUs =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("On-Demand").getRange(5, 6);
  
  switch(action){
      
    case "userJia":
      cUs.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      changeStatus_rColor(status, cUs)
      cUs.setFontColor('white');
      break;
      
    case "removeFromGroup":
      aUs.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      changeStatus_rColor(status, aUs)
      aUs.setFontColor('white');
      break;
      
    case "emailSent":
      eUs.setValue(status).setHorizontalAlignment(DocumentApp.HorizontalAlignment.CENTER);
      changeStatus_rColor(status, eUs)
      eUs.setFontColor('white');
      break;
      
  }
}


function removefromGroups(user_name , group_list , activeSheet , error_1 , error_2 , error_3 , a , username , password ) {

  var b = [];
  
  try{
    for(var group in group_list){
      var groupEmail = group_list[group];
      var group_name = groupEmail[0];
      b.push(group_name)
      
      try {
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
        
      } catch (e) {
        
        Logger.log(e)
        error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
        throw e;
      }
    }
    
  } catch (e) {
    Logger.log(e)
    changeStatus_r("Failed", "removeFromGroup");
    error_2.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    throw e;
    
  }
  
  return true;  
  
} 

function jiraRemove(group_list , user_id , activeSheet , error_1 , error_2 , error_3 , username , password) {
  
  var a = [];
  
  var userJiaStatus = false;
  try{
    
    changeStatus_r("Pending", "userJia");
    changeStatus_r("Pending", "removeFromGroup");
    changeStatus_r("Pending", "emailSent");
    
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
      changeStatus_r("Success", "userJia");
      userJiaStatus = true;
      Logger.log("userJia:success");
    }catch (e) {
      Logger.log(e)
      changeStatus_r("Failed", "userJia");
      error_1.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
      return false;
    } 
    
    Logger.log("userJia:success:process removeFromGroup");
    var removeFromGroupStatus = false;
    if( userJiaStatus ){
      try{
        for(var i in user_id){
          var userId = user_id[i];
          var user_name = userId[0].split("@")[0];
          removefromGroups(user_name , group_list , activeSheet , error_1 , error_2 , error_3 , a  , username , password);
        }
        
        changeStatus_r("Success", "removeFromGroup");
        removeFromGroupStatus = true;
        Logger.log("removeFromGroup:success");
        
      }catch (e) {
        Logger.log(e)
        changeStatus_r("Failed", "removeFromGroup");
        error_1.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
        return false;
      }
      
    }
    
    Logger.log("removeFromGroup:success:process emailNotification");
    if ( removeFromGroupStatus){
      emailNotification(a , group_list , activeSheet , error_1 , error_2 , error_3 , username);  
    }
    
  } catch (e) {
    Logger.log(e)
    error_1.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    return false;
  }   
}

function emailNotification(a, b , activeSheet , error_1 , error_2 , error_3 , username) {
  
  try{
    var email_cc = "rean.GlobalSR@hitachivantara.com"
    var email_to = Session.getActiveUser().getEmail();
    var date = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy");
    var user_count = a.length;
    var group_count = b.length;
    var subject = "[Atlassian Group] Total "+user_count+ " members removed from "+group_count+ " Atlassian groups on "+date ;     
    var atlassianID = username.getResponseText();
    
    var massage = "Total "+user_count+ " members removed from "+group_count+ " Atlassian groups on "+date+". Please refer following table for more details."; 
    
    
    var htmlTemplate = HtmlService
    .createTemplateFromFile("ondemand");
    htmlTemplate.data = [a , b, user_count > group_count ? user_count : group_count , email_to , atlassianID , massage];
    
    var htmlBody = htmlTemplate.evaluate().getContent();
    MailApp.sendEmail({
      to: email_to,
      cc: email_cc,
      subject: subject,
      htmlBody: htmlBody
    }); 
    changeStatus_r("Success", "emailSent");
    
  } catch (e) {
    Logger.log(e.toString()); 
    changeStatus_r("Failed", "emailSent");
    error_3.setValue(e).setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    return false;
    
  }
  
}
