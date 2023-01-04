function COUNT_MEETINGS_BY_SUPPORT_AND_MONTH(dataRange, month, dataRangeTwo) {

    var supportOne = 0
    var supportTwo = 0
    var supportThree = 0
    var supportFour = 0
    var supportFive = 0
    var meetingCount = 0
  
    var ss = SpreadsheetApp.getActive();
    var allsheets = ss.getSheets();
  
    // Acquire number of support per month
    for(var s in allsheets){
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRange)
      var values = range.getValues()
  
      if (sheet.getName().startsWith(month)) { // If sheet matches month
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (col != "" && col.startsWith("Adeel")) {
              supportOne += 1}
            else if (col != "" && col.startsWith("Saad")) {
              supportTwo += 1}
            else if (col != "" && col.startsWith("Chintan")) {
              supportThree += 1}
            else if (col != "" && col.startsWith("Ayman")) {
              supportFour += 1}
            else if (col != "" && col.startsWith("Angela")) {
              supportFive += 1}
          });
        });
      } 
    } 
  
    // Acquire total meetings per month
    for(var s in allsheets) {
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRangeTwo)
      var values = range.getValues()
  
      if (sheet.getName().startsWith(month)) { // If sheet matches month
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (col === "") {} else {meetingCount += 1}
          })
        })
      }
    } 
    supportTotal = supportOne + supportTwo + supportThree + supportFour + supportFive
    
    return meetingCount / supportTotal
  }
  
  
  
  
  
  
  
  
  