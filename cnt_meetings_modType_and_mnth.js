function COUNT_MEETINGS_BY_MOD_TYPE_AND_MONTH(dataRange, month, meetingType, modType) {

  var inCount = 0
  var inCountTwo = 0
  var exCount = 0
  var exCountTwo = 0
  var association = "Association"
  var ab_condo = "AB Condo"

  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();

  var exclude_sheets = ["Moderators","MOps OKRs","Meeting Volume"];
  var internal_mods = ["Amanda Bustard (GQ)", "Nivetha Velupur (GQ)","Kate Polek (GQ)","Rebecca Trotter (GQ)"]
  var external_mods = ["Stacy Costa",	"Sachi Lovatt",	"Chantel Espinola",	"Amy Stephenson",	"Lorena Araujo",	"Joanne Janzen",	"Patricia Walker",	"Rebecca Blankfort",	"Becah Machado-Burton",	"Susan Hwang",	"Sandra Tudge",	"Olivia Dutka",	"Caleb Hayward",	"Cody Sharpe",	"Steph Jacobs",	"Amrinder Singh",	"Wendy Dennis",	"Bhavani Kidambi",	"April Scarff",	"Arshpreet Kaur",	"John Harris",	"Gabriela Jungblut"]

  for(var s in allsheets){
    var sheet = allsheets[s]
    var range = sheet.getRange(dataRange)
    var values = range.getValues()

    if (sheet.getName().startsWith(month)) { // If sheet matches month
      //if (exclude_sheets.includes(sheet.getName())) continue; // Skip sheets in exclude list

        values.forEach(function(row) {
          row.forEach(function(col) {
            if (internal_mods.includes(col) && row[0,5].startsWith(meetingType)) {
              inCount +=1}
            else if (external_mods.includes(col) && row[0,5].startsWith(meetingType)) {
              exCount += 1}
            else if (internal_mods.includes(col) && !row[0,5].startsWith(association) && !row[0,5].startsWith(ab_condo)) {
              inCountTwo += 1}
            else if (external_mods.includes(col) && !row[0,5].startsWith(association) && !row[0,5].startsWith(ab_condo)) {
              exCountTwo += 1}
            else {}
        });
      });
    }
  } // End of Loop

  if (modType.startsWith(meetingType) && modType.startsWith("Internal",28) || modType.startsWith(meetingType) &&modType.startsWith("Internal",31)) {
    return inCount}
  else if (modType.startsWith(meetingType) && modType.startsWith("External",28) || modType.startsWith(meetingType) && modType.startsWith("External",31)) {
    return exCount}
  else if (modType.startsWith(meetingType) && modType.startsWith("Internal", 27)) {
    return inCountTwo}
  else if (modType.startsWith(meetingType) && modType.startsWith("External", 27)) {
    return exCountTwo}
  else if (modType.startsWith(meetingType) && !modType.startsWith(ab_condo) && !modType.startsWith(association)) {
    return inCountTwo + exCountTwo}
  else {
    return inCount + exCount
  }
}