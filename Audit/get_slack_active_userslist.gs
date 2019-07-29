/* 


  ======================================
              Slack Audit 
  ======================================
  
  Written by Nikhil Kapure on 10/04/2018 
  

*/


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('REANCloud IT Operations')
      .addItem('Execute script', 'listAllUsers')
      .addToUi();
} 
  
function listAllUsers() {
  
  var activeSheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  var cell_range = activeSheet.getRange("C2:C");
  var cell_range_2 = activeSheet.getRange("F2:F");
  
  cell_range.clear();
  cell_range_2.clear();
  
  var existingMemberList = [];
  var slack_bot = [];
  var active_count =0;
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var data = scriptProperties.getProperties();
  var token = data.token;
  // Call the Numbers API for random math fact
  var response = UrlFetchApp.fetch("https://slack.com/api/users.list?token="+ token +"&pretty=1");
  
  var fact = response.getContentText();
  var data = JSON.parse(fact);
  
  if (data) {
    var members = data.members;
    for (var i = 0; i < members.length; i++) {
      if (members[i].profile.email){
        if (members[i].deleted){
        } else {
          existingMemberList.push([members[i].profile.email]);
          active_count++;
        };
        
      }   else{
        
        slack_bot.push([members[i].profile.real_name]);
      }
    }
    
  } 
  
  activeSheet.getRange(2, 3, existingMemberList.length).setValues(existingMemberList);
  activeSheet.getRange(2, 6, slack_bot.length).setValues(slack_bot);
  
  Utilities.sleep(60000);// pause in the loop for 60000 milliseconds
  var missing_members = activeSheet.getRange("E2:E").getValues().filter(String);
  
  var active_sub = "[Slack Audit Report] Total active members count is " +active_count+ " suspended from all REAN but active Slack members count is "+missing_members.length;
  if (missing_members.length === 0){
    var response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token="+token+"&channel=CCQ78V9ST&text=%3Athumbsup%3A%20Congratulation!%20We%20are%20good%20here%20for%20Slack%20Tool%20weekly%20audit.%20No%20any%20defect%20found%20in%20this%20week.&as_user=U6ZANQ0V9&pretty=1");
    
    Logger.log("Sucessssssssssssss")
    
  } else{
    
    var response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token="+token+"&channel=CCQ78V9ST&text=%3Ax%3A%20%5BSlack%20Audit%20Report%5D%20Total%20active%20members%20count%20is " +active_count+ " suspended%20from%20all%20REAN%20but%20active%20Slack%20members%20count%20is "+missing_members.length+ "&as_user=U6ZANQ0V9&pretty=1");    
//    sendEmail( missing_members, active_sub, "Found following Missing members in Slack account");
    Logger.log("Faileddddddddddddddd")
    
  }
}

function sendEmail(userList, subject, tagLine){
  try{
    var htmlTemplate = HtmlService
                            .createTemplateFromFile("mail_template");
    htmlTemplate.data = [userList, tagLine];
    var htmlBody = htmlTemplate.evaluate().getContent();
    MailApp.sendEmail({
       to: "GlobalHROps@reancloud.com",
//      to: "nikhil.kapure@reancloud.com",
       cc: "",
       subject: subject,
       htmlBody: htmlBody
    });
  } catch (e) {
      Logger.log(e)
  }
}

