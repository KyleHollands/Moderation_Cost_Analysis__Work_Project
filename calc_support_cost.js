function CALCULATE_SUPPORT_COST(dataRange, month) {

  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();
  var supportCost = 0

  var support_consultants_tierTwo = ["Ayman", "Angela"]
  var support_consultants_tierOne = ["Saad", "Chintan"]

  // Acquire number of support per month
  for(var s in allsheets){
    var sheet = allsheets[s]
    var range = sheet.getRange(dataRange)
    var values = range.getValues()

    // If sheet matches month
    if (sheet.getName().startsWith(month)) { 
      values.forEach(function(row) {
        row.forEach(function(col) {
          if (support_consultants_tierOne.includes(col) || support_consultants_tierOne.includes(row[0,1])) {
            supportCost += 1 * 30 * 6}
          else if (support_consultants_tierTwo.includes(col) || support_consultants_tierTwo.includes(row[0,1])) {
            supportCost += 1 * 30 * 4}
        });
      });
    } 
  } 
  return supportCost 
}
