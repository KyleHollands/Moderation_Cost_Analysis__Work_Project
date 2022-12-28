function COUNT_MODS_BY_MEETING_TYPE(dataRange, modName, meetingType) {

  var count = 0

  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();

  var exclude_sheets = ["Metrics","Moderators"];

  for(var s in allsheets){
    var sheet = allsheets[s];
    var range = sheet.getRange(dataRange)
    var values = range.getValues()

    if (exclude_sheets.includes(sheet.getName())) continue; //Skip sheets in exclude_sheets list
      values.forEach(function(row) {
        row.forEach(function(col) {
          if (col === modName && row[0,5].startsWith(meetingType)) {count +=1}
        });
      });

 // End of loop
  } return count;
}