function formupdate() {
  try{
    
    var existingForm = FormApp.openById('1LZ8SJD-TKYMksu3kuzuSozR5IJoqajsM_Ni52WyCeLQ'); //getting gform ID
    var sectionID = existingForm.getItems();
    Logger.log(sectionID)
    var sectionManagerID = sectionID[125].getId();
    var activeSheet = SpreadsheetApp.getActive().getSheetByName("Master_Data"); //Selecting sheet
    var managerID = activeSheet.getRange("I:I").getValues().filter(String); //getting all active and service notice period employee list
    
    var employeeItems = existingForm.getItemById(sectionManagerID);
    
    Logger.log("Updating Manager Email IDs")
    
    var item = employeeItems.asListItem();
    item.setTitle('Manager email ID').setChoiceValues(managerID)
    item.setHelpText("Kindly select your manager Email ID from following drop-down list")
    
    Logger.log("Updated Manager Email IDs")   
    
  }catch (e){
    Logger.log(e); 
    
  }
}
