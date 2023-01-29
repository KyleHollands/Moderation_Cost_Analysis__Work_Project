function COUNT_MODS_BY_MEETING_TYPE(dataRange, modName, meetingType) {

  var count = 0
  var countTwo = 0
  var countThree = 0

  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();

  var exclude_sheets = ["Moderators","MOps OKRs","Meeting Volume"];

  for(var s in allsheets){
    var sheet = allsheets[s];
    var range = sheet.getRange(dataRange)
    var values = range.getValues()

    //Skip sheets in exclude_sheets list
    if (exclude_sheets.includes(sheet.getName())) continue; 
      values.forEach(function(row) {
        row.forEach(function(col) {
          if (col === modName && row[0,5].startsWith("AB Condo")) {count +=1}
          else if (col === modName && row[0,5].startsWith("Association")) {countThree += 1}
          else if (col === modName && !row[0,5].startsWith("Association") && !row[0,5].startsWith("AB Condo")) {countTwo += 1} 
          else {}
        });
      });
      
 // End of loop
  } if (meetingType.startsWith("Ontario")) {return countTwo}
    else if (meetingType.startsWith("AB Condo")) {return count}
    else if (meetingType.startsWith("Association")) {return countThree}
    else {}
}
