function main() {
  try{
    
    var userEmail = Session.getActiveUser().getEmail();
    Logger.log("Executed by: "+userEmail);
    
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atlassian Onboarding");
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
    
    
    //      var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("testing");
    var table = activeSheet.getRange(8,2,16,2).getValues();
    var primaryEmailAddress = activeSheet.getRange(4,2).getValue();
    
    cell_range.clear();
    
    cell_range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    cell_range.setFontFamily("Cambria");
    cell_range.setFontSize(12);
    
    var createUserStatus =  activeSheet.getRange(9, 6);
    var addUserToGroupStatus =  activeSheet.getRange(10, 6);
    var emailsentStatus =  activeSheet.getRange(11, 6);
    
    
    createUserStatus.setValue("User Account Create").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    addUserToGroupStatus.setValue("Add User in Groups").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    emailsentStatus.setValue("Email Sent").setHorizontalAlignment(DocumentApp.HorizontalAlignment.LEFT);
    
    setStatus("Pending", "createUser");
    setStatus("Pending", "adduserToGroup");
    setStatus("Pending", "emailsentonStatus");
    
    HR_TEAM = activeSheet.getRange(4,6).getValue();
    HR_ID = activeSheet.getRange(5,6).getValue();
    
    var user = addInJira(primaryEmailAddress, table[1][1],table[2][1],table[4][1], table[5][1], table[6][1], table[8][1], table[9][1], table[14][1], table[15][1], HR_TEAM, HR_ID, table[10][1] , username , password);
  } catch (e) {
    Logger.log(e);
  }
}
