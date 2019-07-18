function newGroup(groupID, groupName, groupDescription, groupMemberList, groupOwnerList) {
  
  var getGroupStatus = false;
  
  try{
    // Get Group
    var  user = AdminDirectory.Groups.get(groupID);
    Logger.log(user.adminCreated)
    var getGroupStatus = true;
    
  } catch(e){
    Logger.log(e)
  }    
  
  
  if (getGroupStatus){
    Logger.log("Group Found")
    
  }
  else {
    
    Logger.log("Group Not Found. Creating a Group "+groupID)
    
    var groupDetails = {
      email: groupID,
      description: groupDescription,
      name: groupName
    };
    
    var  groupCreation = AdminDirectory.Groups.insert(groupDetails);
    Logger.log("Group Created Sucessfully")    
  }
}

function addingMembersInGroup(groupID, groupMemberList){
  
  for (i in groupMemberList){
    
    var member = groupMemberList[i];
    var domain = member.split("@")[1];
    
    if (domain === "reancloud.com"){
      
      var  addInGroupMember = AdminDirectory.Members.hasMember(groupID, member);
      
      if (addInGroupMember.isMember){
        
        Logger.log("Part of this group")
        
      } else {
        
        Logger.log("Not a member of this group")
        
        var role = {
          email: member,
          role: 'MEMBER'
        };
        
        var  addAsAMember = AdminDirectory.Members.insert(role, groupID);
        Logger.log("Members added in "+groupID+" group sucessfully")
      }
    }
    try {
      
      var role = {
        email: member,
        role: 'MEMBER'
      };
      var  addAsAMember = AdminDirectory.Members.insert(role, groupID);
      
    }catch(e){
      Logger.log(e)
    }    
  }

}

function groupOwnerChange(groupID, groupOwnerList){

  Utilities.sleep(30000); // pause in the loop for 30000 milliseconds
    if (groupOwnerList) {
      for (i in groupOwnerList){
        var memberOwner = groupOwnerList[i];      
        var  addInGroupOwner = AdminDirectory.Members.hasMember(groupID, memberOwner);
        Logger.log(addInGroupOwner)
        var member  = {
          email: memberOwner,
          role: 'OWNER'          
        };
        
        if (addInGroupOwner.isMember){
          
          Logger.log("Updating Role")          
          var addAsAOwner = AdminDirectory.Members.update(member, groupID, memberOwner);
          Logger.log("Updated Role")          
          
        } else {
          
          Logger.log("Adding as a Owner")
          var addAsAOwner = AdminDirectory.Members.insert(member, groupID);
          Logger.log("Added as a Owner")
          
        }
      }
      
    }
    
    else 
    {
      Logger.log("Owner not provided")
    }  

}

function memberDelete(groupID, deleteMembers){
  
  for (i in deleteMembers){
    
    var member = deleteMembers[i]; 
    var domain = member.split("@")[1];
    
    if (domain === "reancloud.com"){
      
      var  memberStatus = AdminDirectory.Members.hasMember(groupID, member);
      
      if (memberStatus.isMember){
        
        var deleteFrouGroup = AdminDirectory.Members.remove(groupID, member); 
        Logger.log("User has been removed from"+groupID+" group")
      }  
    } else {
      
      try {
        
        var deleteFrouGroup = AdminDirectory.Members.remove(groupID, member); 
        Logger.log("User has been removed from"+groupID+" group")
        
      }catch(e){
        Logger.log(e)
      }  
    }
  }
}

function groupDelete(groupID){
  
  var getGroupStatus = false;
  
  try{
    
    // Get Group
    var  user = AdminDirectory.Groups.get(groupID);
    Logger.log(user.adminCreated)
    var getGroupStatus = true;
    
  } catch(e){
    Logger.log(e)
  }    
  
  if (getGroupStatus){
    
    Logger.log("Group Found")    
    var  groupCreation = AdminDirectory.Groups.remove(groupID);
    Logger.log("Removed Group "+groupID+" Sucessfully")
  }
  else {  
    Logger.log("Group Not Found")
  }
  
}  


function updateGroupName(groupID, groupName) {
  
  var group = AdminGroupsSettings.newGroups();
  group.name = groupName,
    
    AdminGroupsSettings.Groups.patch(group, groupID);
  
  Logger.log("Changed Group Name.");
  
}

function updateGroupDescription(groupID, groupDescription){
  
    var group = AdminGroupsSettings.newGroups();
  group.description = groupDescription,
    
    AdminGroupsSettings.Groups.patch(group, groupID);
  
  Logger.log("Changed Group description.");
  
}


function publicGroup(groupID) {
  
  var group = AdminGroupsSettings.newGroups();
  group.messageDisplayFont = "DEFAULT_FONT";
  group.whoCanInvite = "ALL_MANAGERS_CAN_INVITE";
  group.allowWebPosting = "true";
  group.allowExternalMembers = "false";
  group.sendMessageDenyNotification = "false";
  group.kind = "groupsSettings#groups";
  group.membersCanPostAsTheGroup = "false";
  group.customFooterText = "";
  group.whoCanPostMessage = "ANYONE_CAN_POST";
  group.whoCanContactOwner = "ANYONE_CAN_CONTACT";
  group.whoCanViewGroup = "ALL_IN_DOMAIN_CAN_VIEW";
  group.customReplyTo = "";
  group.replyTo = "REPLY_TO_SENDER";
  group.spamModerationLevel = "ALLOW";
  group.archiveOnly = "false";
  group.allowGoogleCommunication = "false";
  group.maxMessageBytes = 26214400;
  group.isArchived = "true";
  group.defaultMessageDenyNotificationText = "";
  group.showInGroupDirectory = "true";
  group.includeInGlobalAddressList = "true";
  group.whoCanViewMembership = "ALL_IN_DOMAIN_CAN_VIEW";
  group.messageModerationLevel = "MODERATE_NONE";
  group.whoCanLeaveGroup = "ALL_MEMBERS_CAN_LEAVE";
  group.includeCustomFooter = "false";
  group.whoCanJoin = "CAN_REQUEST_TO_JOIN";
  group.whoCanAdd = "ALL_MANAGERS_CAN_ADD"
 
  AdminGroupsSettings.Groups.patch(group, groupID);

  Logger.log("Public Group setting applied sucessfully")
}

function privateGroup(groupID) {
  var group = AdminGroupsSettings.newGroups();
  
  group.messageDisplayFont = "DEFAULT_FONT";
  group.whoCanInvite = "ALL_MANAGERS_CAN_INVITE";
  group.allowWebPosting = "true";
  group.allowExternalMembers = "false";
  group.sendMessageDenyNotification = "false";
  group.kind = "groupsSettings#groups";
  group.membersCanPostAsTheGroup = "false";
  group.customFooterText = "";
  group.whoCanPostMessage = "ALL_IN_DOMAIN_CAN_POST";
  group.whoCanContactOwner = "ANYONE_CAN_CONTACT";
  group.whoCanViewGroup = "ALL_IN_DOMAIN_CAN_VIEW";
  group.customReplyTo = "";
  group.replyTo = "REPLY_TO_IGNORE";
  group.spamModerationLevel = "MODERATE";
  group.archiveOnly = "false";
  group.allowGoogleCommunication = "false";
  group.maxMessageBytes = 26214400;
  group.isArchived = "true";
  group.defaultMessageDenyNotificationText = "";
  group.showInGroupDirectory = "false";
  group.includeInGlobalAddressList = "true";
  group.whoCanViewMembership = "ALL_IN_DOMAIN_CAN_VIEW";
  group.messageModerationLevel = "MODERATE_NONE";
  group.whoCanLeaveGroup = "ALL_MEMBERS_CAN_LEAVE";
  group.includeCustomFooter = "false";
  group.whoCanJoin = "CAN_REQUEST_TO_JOIN";
  group.whoCanAdd = "ALL_MANAGERS_CAN_ADD"  
  
  AdminGroupsSettings.Groups.patch(group, groupID);
  
  Logger.log("Private Group setting applied sucessfully");
}

function sendEmailAuto(key, value, subject, message, global_support_id, requested){
  Logger.log("+++++++++++++++++++++++++++++++++++++++++++++")
  var htmlTemplate = HtmlService
  .createTemplateFromFile("template");
  htmlTemplate.data = [key,value, message];
  
  var htmlBody = htmlTemplate.evaluate().getContent();
  
  MailApp.sendEmail({
    to: requested,
    cc: global_support_id,
    //    cc: "nikhil.kapure@reancloud.com",
    subject: subject,
    htmlBody: htmlBody
  });
   Logger.log("_________________________________________")
 
}
