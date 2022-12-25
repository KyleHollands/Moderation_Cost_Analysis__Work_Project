function LOOP_THROUGH_SHEETS_FOUR(rangeOne, countOne, countTwo) {

  var count = 0

  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();

  var exclude = ["Meeting Volume","Moderators"];

  for(var s in allsheets){
    var sheet = allsheets[s];
    var range = sheet.getRange(rangeOne)
    var values = range.getValues()

    // Stop iteration execution if the condition is met.
    //-------------------------------------------------
    if(exclude.indexOf(sheet.getName()) == 1) continue;
    //-------------------------------------------------
      values.forEach(function(row) {
        row.forEach(function(col) {
          if (col === countOne && row[0,5] === countTwo) {count +=1}
        });
      });
 // End of loop
  } return count;
}