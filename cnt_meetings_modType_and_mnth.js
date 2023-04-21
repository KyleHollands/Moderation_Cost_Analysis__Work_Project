function COUNT_MEETINGS_BY_MOD_TYPE_AND_MONTH(dataRange, month, meetingType) {

  var inCount = 0
  var inCountTwo = 0
  var inCountThree = 0
  var exCount = 0
  var exCountTwo = 0
  var exCountThree = 0

  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();

  //var exclude_sheets = ["Moderators","MOps OKRs","Meeting Volume"];
  var internal_mods = ["Amanda Bustard (GQ)", "Nivetha Velupur (GQ)","Kate Polek (GQ)","Rebecca Trotter (GQ)"]
  var external_mods = ["Stacy Costa",	"Sachi Lovatt",	"Chantel Espinola",	"Amy Stephenson",	"Lorena Araujo",	"Joanne Janzen",	"Patricia Walker",	"Rebecca Blankfort",	"Becah Machado-Burton",	"Susan Hwang",	"Sandra Tudge",	"Olivia Dutka",	"Caleb Hayward",	"Cody Sharpe",	"Steph Jacobs",	"Amrinder Singh",	"Wendy Dennis",	"Bhavani Kidambi",	"April Scarff",	"Arshpreet Kaur",	"John Harris",	"Gabriela Jungblut","Eric Aspila","Kathleen Speckert","Liona Davies","Eric Szyiko","Andrew Natale","Sheila McFadyen","Leanne Unruh","Sharry Ash","Jaspreet Singh","Eric Gloutnez","Justin Trent","Ren Haskett","Andreane Dicaire","Vidhi Abbi","Fatima Mehrzad","Elisa Brosseau"]

  for(var s in allsheets){
    var sheet = allsheets[s]
    var range = sheet.getRange(dataRange)
    var values = range.getValues()

    // If sheet matches month
    if (sheet.getName().startsWith(month)) { 

      // Skip sheets in exclude list
      //if (exclude_sheets.includes(sheet.getName())) continue;

        values.forEach(function(row) {
          row.forEach(function(col) {
            if (internal_mods.includes(col) && row[0,5].startsWith("AB Condo")) {inCount +=1}
            else if (external_mods.includes(col) && row[0,5].startsWith("AB Condo")) {exCount += 1}
            else if (internal_mods.includes(col) && row[0,5].startsWith("Association")) {inCountThree += 1}
            else if (external_mods.includes(col) && row[0,5].startsWith("Association")) {exCountThree += 1}
            else if (internal_mods.includes(col) && !row[0,5].startsWith("Association") && !row[0,5].startsWith("AB Condo")) {inCountTwo += 1}
            else if (external_mods.includes(col) && !row[0,5].startsWith("Association") && !row[0,5].startsWith("AB Condo")) {exCountTwo += 1}
            else {}
        });
      });
    }
  } // End of Loop

  if (meetingType.startsWith("AB Condo") && meetingType.startsWith("Internal",28)) {return inCount}
  else if (meetingType.startsWith("Association") && meetingType.startsWith("Internal", 31)) {return inCountThree}
  else if (meetingType.startsWith("AB Condo") && meetingType.startsWith("External",28)) {return exCount}
  else if (meetingType.startsWith("Association") && meetingType.startsWith("External", 31)) {return exCountThree}
  else if (meetingType.startsWith("Ontario") && meetingType.startsWith("Internal", 27)) {return inCountTwo}
  else if (meetingType.startsWith("Ontario") && meetingType.startsWith("External", 27)) {return exCountTwo}
  else if (meetingType.startsWith("Ontario") && !meetingType.startsWith("AB Condo") && !meetingType.startsWith("Association")) {return inCountTwo + exCountTwo}
  else if (!meetingType.startsWith("Ontario") && !meetingType.startsWith("AB Condo")) {return inCountThree + exCountThree}
  else if (!meetingType.startsWith("Ontario") && !meetingType.startsWith("Association")) {return inCount + exCount}
  else {}
}