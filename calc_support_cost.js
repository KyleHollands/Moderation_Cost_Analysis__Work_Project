function CALCULATE_SUPPORT_COST(dataRange, month) {

    var ss = SpreadsheetApp.getActive();
    var allsheets = ss.getSheets();
    var supportCost = 0
  
    var support_consultants = ["Saad", "Chintan", "Ayman", "Angela"]
  
    // Acquire number of support per month
    for(var s in allsheets){
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRange)
      var values = range.getValues()
  
      // If sheet matches month
      if (sheet.getName().startsWith(month)) { 
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (support_consultants.includes(col) || support_consultants.includes(row[0,1])) {
              supportCost += 1 * 30 * 5}
          });
        });
      } 
    } 
    return supportCost 
  }
  