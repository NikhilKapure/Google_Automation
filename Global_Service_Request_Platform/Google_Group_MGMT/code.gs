/* 


  ======================================
          Google Groups Audit 
  ======================================
  
  Written by Nikhil Kapure on 23/01/2018 
  

*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('REANCloud IT Operations')
      .addItem('Execute script', 'listAllGroups')
        .addItem('Final List', 'main')
      .addToUi();
} 

function listAllGroups() {

  // The code below logs the index of a sheet named "Sheet1"
  //var activeSheet = SpreadsheetApp.getActiveSheet().getSheetByName("Sheet1");
 var activeSheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
// var excludeUsersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exclude Group");
// var excludeUsers = excludeUsersSheet.getRange(2, 1,excludeUsersSheet.getLastRow()).getValues();
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
 if (activeSheet != null) {
   Logger.log("Index: "+activeSheet.getIndex());
 }
  var header = ["Group Name", "Group Id", "Members", "Role", "User Status"];
  var rows = [];
  
    var cell_range = activeSheet.getRange("A:E");
  activeSheet.clear();
  
//  var menucell = activeSheet.getRange("F1:H4");
//  activeSheet.getRange(1, 6, 3, 3).setBackground('#08009A');
  activeSheet.appendRow(header);
  activeSheet.getRange(1,1,1,5).setBackground('#f4a142');
  
  var pageTokenGroup, pageGroup, pageTokenMembers, pageMemebers; 
  var rows = [];
  var suspended_count = 0;
  
  do {
    pageGroup = AdminDirectory.Groups.list({
      domain: 'reancloud.com',
      maxResults: 500,
      pageToken: pageTokenGroup
    });
    
    var groups = pageGroup.groups;
    if (groups) {
      for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        var skip = true;
//        Logger.log(excludeUsers)
//        for( i in excludeUsers)
//          if(excludeUsers[i][0]===group.name ){
//            Logger.log("Skipping "+group.name);
//            skip = true;
//         }
       if(skip){
//          Logger.log('%s (%s)', group.name, group.email);
//          var row = [group.name, group.email, "", "", ""];
//          rows.push(row);
           
          do {
             pageMemebers = AdminDirectory.Members.list(group.email,{
             domainName: 'reancloud.com',
             maxResults: 500,
             pageToken: pageTokenMembers,
           });
           var members = pageMemebers.members
           if (members)
           {
             for (var j = 0; j < members.length; j++)
             {
               var member = members[j];
               if (member.status === "SUSPENDED") {
                 
                 try{
                 Logger.log(group.email, member.email)              
                 memberDelete(group.email, member.email);
                 } catch (e){
                 Logger.log(e)
                 }
               } else {
//               var row = ["", "", member.email, member.role, member.status];
                 var row = [group.name, group.email, member.email, member.role, member.status];
//               Logger.log(row)
               rows.push(row);
               }
             }
           }
           pageTokenMembers = pageMemebers.nextPageToken;
         } while (pageTokenMembers);
       }
      }
    } else {
      Logger.log('No groups found.');
    }
    pageTokenGroup = pageGroup.nextPageToken;
  } while (pageTokenGroup);
  
  activeSheet.getRange(activeSheet.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
  
  var maxRows = activeSheet.getMaxRows();
  var values = activeSheet.getRange(2, 3, maxRows-1, 3).getValues();

  for(var rowNum = 0; rowNum < values.length; rowNum++){
    var CRange = activeSheet.getRange(rowNum+2, 3);
   if (values[rowNum][0] == ""){
     CRange.setBackground("#070000");
    }

    else if (values[rowNum][0].indexOf("reancloud.com") == -1){
      CRange.setBackground("#1100FF");
      CRange.setFontColor('#FEFEFE');
    }
    else if(rowNum != 0 ){
      CRange.setBackground("#ffffff");
    }
    var DRange = activeSheet.getRange(rowNum+2, 4);
    if (values[rowNum][1] == "MEMBER"){
      DRange.setBackground("#FF00FB");
      DRange.setFontColor('#FEFEFE');
    }
     else if (values[rowNum][1] == "OWNER"){
      DRange.setBackground("#FF7D03");
      DRange.setFontColor('#050505');
    }
     else if (values[rowNum][1] == "MANAGER"){
      DRange.setBackground("#E5FF00");
      DRange.setFontColor('#050505');
    }
     else if (values[rowNum][1] == ""){
      DRange.setBackground("#ffffff");
    }

    var ERange = activeSheet.getRange(rowNum+2, 5);
    if ( values[rowNum][2] == "ACTIVE"){
      ERange.setBackground("#135b03");
      ERange.setFontColor('#FEFEFE');
    }
    else if (values[rowNum][2] == "SUSPENDED"){
      suspended_count++;
      ERange.setBackground("#db2323");
      ERange.setFontColor('#FEFEFE');
    }
    else if (values[rowNum][2] == ""){
      ERange.setBackground("#ffffff");
    }
       
  }
   
  var allCells = activeSheet.getRange("A1:E5000");
  activeSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  allCells.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  allCells.setFontFamily("Cambria");
  allCells.setFontSize(11);

}

function memberDelete(groupID, memberID){

  var  memberStatus = AdminDirectory.Members.hasMember(groupID, memberID);

  try{
    
    var deleteFrouGroup = AdminDirectory.Members.remove(groupID, memberID); 
    Logger.log("User has been removed from"+groupID+" group")
  }  catch(e){
    Logger.log(e)
  }
  
}
