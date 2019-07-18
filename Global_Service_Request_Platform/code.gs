/* 


  ======================================
         Email Notification  
  ======================================
  
  Written by Nikhil Kapure on 03/07/2018 
  

*/

function SendGoogleForm(){  
  try 
  {  
    Logger.log("Started");
    var email_to = "";
    var email_cc = "";
    var ticket_id = "rean.GlobalSR@hitachivantara.com";
    var global_support_id = "GlobalHROps@reancloud.com";
    var mgs_support = "ms@reancloud.com"
    var jira_support = "jira-admin@reancloud.com,globalsr@reancloud.com";
    var tempo_admin = "tempo-admin@reancloud.com";
    var elizabeth_ID = "elizabeth.ford@reancloud.com";
    var rupa_ID = "rupa@reancloud.com";
    
    var s = SpreadsheetApp.getActive().getSheetByName("Form Responses 1");
    var columns = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
    var lastRow = s.getLastRow();
    var row = s.getRange(lastRow,1,1,s.getLastColumn()).getValues()[0];  
    var manager = row[3].split("@", 2)[0];
    var manager_id = row[3];
    var priority = row[5];
    var country = row[6];
    var expense = row[10];
    var timeline = "";
    var it_tool = row[7];
    var requested = row[2];
    var requesterName = row[2].split("@", 2)[0];
    var date = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy");
    var status = true;    
    
    var MasterData = SpreadsheetApp.getActive().getSheetByName("Master_Data"); //Selecting sheet
    var directorIDs = MasterData.getRange("O2:O").getDisplayValues().filter(String); //getting all active and service notice period employee list

    for ( var i in directorIDs ) {
      
      var directorID = directorIDs[i][0];
     
      if ( requested === directorID ){
        
        Logger.log("Requester is Director and Above.")
        
        var key = [];
        var value = [];        
        for ( var keys in columns ) {
          if( keys != ""){
            if ( row[keys] && (row[keys] != "") ) {
              Logger.log(row[keys])
              key.push(columns[keys]); 
              value.push(row[keys]);
            }
          }
        }
        
        var groupID = row[59];
        var groupOperation = row[60];
        var emailOperation = row[71];
        
        if ( groupID != "" || emailOperation === "New Creation" ) {
          
          var status = false;
          
          if ( groupID != "" ) {
            
            if (groupOperation === "New Creation"){
              
              Logger.log("Performing "+groupOperation+" Operation");
              
              var groupPrivacy = row[61];    
              var groupName = row[62];
              var groupDescription = row[63];
              var groupMemberList = row[64].split(',');
              var groupOwnerList = row[65].split(',');
              
              newGroup(groupID, groupName, groupDescription, groupMemberList, groupOwnerList);
              addingMembersInGroup(groupID, groupMemberList);
              groupOwnerChange(groupID, groupOwnerList);
              groupOwnerChange(groupID, groupOwnerList);
              
              if ( groupPrivacy === "Private") {
                
                Logger.log("Updating Private access level "+groupID+" Google Group");
                privateGroup(groupID);
                
              } else if ( groupPrivacy === "Public") {
                
                Logger.log("Updating Public access level "+groupID+" Google Group");
                publicGroup(groupID);
                
              }     
              Logger.log("Calling Email Notify function")
              
              var subject = "[Google Group Management] "+requested+ " requested for Group "+groupOperation+" on "+date; 
              var message = "New Group "+groupID+" has been created Sucessfully";     
              
              sendEmailAuto(key, value, subject, message, global_support_id, requested); 
              s.getRange(lastRow, 1).setValue("Auto Resolved");

            } //New group creation loop    
            
            else if ( groupOperation === "Modification"){
              
              Logger.log("Performing "+groupOperation+" Operation");
              
              var groupMemberList = row[69].split(',');
              var groupOwnerList = row[70].split(',');
              var groupPrivacy = row[68]
              
              Logger.log(groupOwnerList)
              
              if (groupMemberList != "") {
                
                Logger.log("Adding members in Group "+groupID)
                addingMembersInGroup(groupID, groupMemberList);
              }    
              
              if (groupOwnerList != "") {
                
                Logger.log("Updating Owner in Group "+groupID)
                groupOwnerChange(groupID, groupOwnerList);
                
              }     
              
              if (groupPrivacy === "Private") {
                
                Logger.log("Updating Private access level "+groupID+" Google Group");
                privateGroup(groupID);
                
              } else if (groupPrivacy === "Public"){
                
                Logger.log("Updating Public access level "+groupID+" Google Group");
                publicGroup(groupID);      
              } 
              
              var subject = "[Google Group Management] "+requested+ " requested for Group "+groupOperation+ " on "+date; 
              var message = groupID+" Sucessfully Modify";     
              
              sendEmailAuto(key, value, subject, message, global_support_id,requested);
              s.getRange(lastRow, 1).setValue("Auto Resolved");             
              
            }// group modification loop    
            
            else if ( groupOperation === "Deletion"){
              
              var deleteOperation = row[66];
              
              if ( deleteOperation === "Group Delete") {
                
                Logger.log("Deleting "+groupID+" Google Group")
                groupDelete(groupID);
                
                var subject = "[Google Group Management] "+requested+ " requested for Group "+groupOperation+ " on "+date; 
                var message = groupID+" Deleted Sucessfully";     
                
                sendEmailAuto(key, value, subject, message, global_support_id,requested);
                s.getRange(lastRow, 1).setValue("Auto Resolved");                
                
              } else {
                
                var groupDeletionMemberList = row[67].split(',');
                Logger.log("Deleting members from "+groupID+" Google Group")
                var subject = "[Google Group Management] "+requested+ " requested for Group "+groupOperation+ " on "+date; 
                var message = groupDeletionMemberList+"members from "+groupID;
                memberDelete(groupID, groupDeletionMemberList);
                sendEmailAuto(key, value, subject, message, global_support_id, requested);
                s.getRange(lastRow, 1).setValue("Auto Resolved");                
            
              } 
              
            } //Group Deletion loop.
            
          } // group level operation.
          
          if ( emailOperation === "New Creation"){
            
            try{
              
              var firstName = row[72];
              var lastName = row[73];                
              var primaryEmailAddress = row[74];
              
              var pass = Math.random().toString(36);
              var user = {
                primaryEmail: primaryEmailAddress,
                name: {
                  givenName: firstName,
                  familyName: lastName
                },
                password: pass
              };
              user = AdminDirectory.Users.insert(user);
              
              var message = "Please refer newly created account details = username: "+user.primaryEmail+" password: "+pass;                      
                      var subject = "[Auto Resolved SR] "+requesterName+" requested for New Email ID Creation on "+date;        
              sendEmailAuto(key, value, subject, message, global_support_id,requested);
              s.getRange(lastRow, 1).setValue("Auto Resolved");              
              
              Logger.log('User %s created with ID %s.', user.primaryEmail, user.id);
              
              var status = false;
              
            } catch (e){
              Logger.log(e)
            }
            
          } // New email Creation 
        } // group or email ID operation
        
      } // selected director loop
      if (status){
        var status = true;      
      }
      else{
      }
      
    } // Director loop 
    
    if (status){
      
      //Base on country selecting support ID and preparing subject    
      if ( country === "India" ){
        
        var support_id = "GlobalHROps@reancloud.com";
        var subject = "[India Internal SR] - "+row[4]+" - Priority "+priority+" - requested by - "+row[2]+" on " +date;  
        
      }
      else {
        
        var support_id = "GlobalHROps@reancloud.com";
        var subject = "[US Internal SR] - "+row[4]+" - Priority "+priority+" - requested by - "+row[2]+" on " +date;  
        
      }
      
      // redirect to tool owners.
      if( it_tool === "Jira Access" || it_tool === "Jira Issue") {
        
        var ticket_id = jira_support;
        var support_id = "";
      }
      else if( it_tool === "GitHub" || it_tool === "AWS" || it_tool === "Google Cloud Account" || it_tool === "Datadog" || it_tool === "CMP (Cloud Management Platform)" || it_tool === "CloudHealth" || it_tool === "Wormly tool" || it_tool === "MS Azure Account") {
        
        var ticket_id = mgs_support;
        var support_id = "";
      }
      else if( it_tool === "Tempo/Time Log issues") {
        
        var ticket_id = tempo_admin;
        var support_id = "";
      }
      else if (row[49] === "Insider Threat/Confidentiality Breach" ){
        
        var ticket_id = elizabeth_ID;
        var support_id = rupa_ID;
        
      }
      else if ( it_tool === "New Additional Software/License" || it_tool === "Adobe License" || it_tool === "Code School" || it_tool === "Confluence Access Request/Update/Change" || it_tool === "Team Drive" || it_tool === "Google Groups" || it_tool === "Google Email" || it_tool === "Google Calendar" || it_tool === "LastPass" || it_tool === "Linux Academy/ACloudGuru/Safari" || it_tool === "Lucidchart" || it_tool === "Microsoft License" || it_tool === "Slack issue" || it_tool === "Slack Channel Integration" || it_tool === "Slack Public/Private Channel" || it_tool === "Shutterstock" ){
        
        var ticket_id = "globalsr@reancloud.com";
        var support_id = global_support_id;
      }
      else{
      }  
      
      // Set priority
      if( priority === 1 ) {
        timeline = "24";
        Logger.log("24 HRS");
      }
      else if( priority === 2) {
        timeline = "48";
        Logger.log("48 HRS");
        
      }
      else if( priority === 3 ) {
        timeline = "72";
        Logger.log("72 HRS");
        
      }  else{
        Logger.log("No match");
        
      } 
      
      if( row[3] === ""){
        email_to = ticket_id +","+ support_id;
        email_cc = row[2];
        var message = "<p>Hi Team,\n\n <p>Please look into the below request.</p>" + "\n\n";
        //      message += "<p><b>TAT is for this " + timeline  + " hours. </b></p>" + "\n\n";
        
        message += "=====================================================================================\n\n";
      }
      else{
        email_to = row[3] +","+ ticket_id;
        email_cc = row[2] +","+ support_id;
        var message = "Hello " + manager+"," + "\n\n";
        message += "<p>Please look into the below request. <b> Manager (" + manager +")</b> reply to this email with your approval or disapproval.</p>" + "\n\n";
        message += "<p><b>Please have this ticket approved within 24 hours by your direct manager, or it will automatically be closed.</b></p>" + "\n\n";
        message += "<p>Support Team:</p>" + "\n\n";
        message += "<p>Please do not act on this ticket unless it is approved by " + manager_id  + "</p>" + "\n\n";
        //      message += "<p><b>TAT is for this " + timeline  + " hours after manager approval.</b></p>" + "\n\n";      
        message += "=====================================================================================\n\n";
      }
      
      // Auto response AV request
      if( it_tool === "Trend Micro Antivirus License"){
        message += "<p>If you want to <b>install AV on your machine</b> then please use following url - </p>"  + "\n";
        message += "<p>http://wfbs-svc-nabu.trendmicro.com/wfbs-svc/download/en/view/activation_mgclink?id=8We%2BACwsxMnmRLFvM0cSQjRyWMq0a0IA79Yyv15Qh2ZSgA%3D%3D</p>" + "\n\n";
        message += "=====================================================================================\n\n";
      }            
      else{
      }
      
      // Auto response Expense Reimbursement request
      if( expense === "Expense Reimbursement"){
        message += "<p>Kindly find below the link to claim reimbursements, kindly fill the same.</p>"  + "\n";
        message += "<p>https://docs.google.com/forms/d/1SfT6_cQBmG97rjhs2m1VCYUiZF8z9npqmAkTjMnMpWw/viewform?ts=5a852232&edit_requested=true</p>" + "\n\n";
        message += "=====================================================================================\n\n";
      }            
      else{
      }    
      
      
      for ( var keys in columns ) {
        if( keys != ""){
          if ( row[keys] && (row[keys] != "") ) {
            Logger.log(row[keys])
            message += "<p><strong>"+columns[keys]+' :: '+"</strong>"+ row[keys] + "</p>" + "\n\n"; 
          }
        }
      }
      
      message += "=====================================================================================\n";
      message += "<p>Regards,</p>";
      message += "<p>REANCloud Operations Team</p>";
      
      MailApp.sendEmail({
        to: email_to,
        cc: email_cc,
        subject: subject,
        htmlBody: message
      });    
      
      s.getRange(lastRow, 1).setValue("Email Sent");
    }
    else{
      Logger.log("This issue has been auto resolved.")      }
    
  } catch (e) {
    Logger.log(e.toString());
  }
  
}

