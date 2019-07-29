function main() {
  
  var activeSheet = SpreadsheetApp.getActive().getSheetByName("Main");
  var dataSheet = SpreadsheetApp.getActive().getSheetByName("Formulas");
  
  activeSheet.clear();
  var groupList = dataSheet.getRange("C2:C").getValues().filter(String);
  var header = ["Group Name", "Group Id", "Members", "Role", "User Status", "Group Privacy"];
  activeSheet.appendRow(header);
  activeSheet.getRange(1,1,1,6).setBackground('#f4a142');
  var rows = []
  
  for (i in groupList){
    
    var groupID = groupList[i];
    var  group = AdminDirectory.Groups.get(groupID);
  var data = AdminGroupsSettings.Groups.get(groupID);
    
    if ( data.whoCanPostMessage === "ANYONE_CAN_POST"){
      Logger.log("Public")
      var row = [group.name, group.email, "", "", "","PUBLIC"];
      
    }else {
      Logger.log("Private")
      var row = [group.name, group.email, "", "", "","PRIVATE"];
      
    }
    
//    var row = [group.name, group.email, "", "", ""];
    rows.push(row);
    
    var  responce = AdminDirectory.Members.list(group.email);
    var membersList = responce.members
    if (membersList)
    {
      for (var j = 0; j < membersList.length; j++)
      {
        var member = membersList[j];
        
        var row = ["", "", member.email, member.role, member.status,""];
        rows.push(row);
      }
    }
    
  }
  
  activeSheet.getRange(activeSheet.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
  
  
  
  var maxRows = activeSheet.getMaxRows();
  var values = activeSheet.getRange(2, 3, maxRows-1, 4).getValues();
  
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
      
      ERange.setBackground("#db2323");
      ERange.setFontColor('#FEFEFE');
    }
    else if (values[rowNum][2] == ""){
      ERange.setBackground("#ffffff");
    }

    var FRange = activeSheet.getRange(rowNum+2, 6);
    if ( values[rowNum][3] == "PRIVATE"){
      FRange.setBackground("#135b03");
      FRange.setFontColor('#FEFEFE');
    }
    else if (values[rowNum][3] == "PUBLIC"){
      
      FRange.setBackground("#db2323");
      FRange.setFontColor('#FEFEFE');
   
  }
  }  
  var allCells = activeSheet.getRange("A1:F5000");
  activeSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  allCells.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  allCells.setFontFamily("Cambria");
  allCells.setFontSize(11);
}

