function COUNT_MEETINGS_BY_MOD_TYPE_AND_MONTH(dataRange, month, meetingType, modType) {

    var inCount = 0
    var inCountTwo = 0
    var exCount = 0
    var exCountTwo = 0
    var association = "Association"
    var ab_condo = "AB Condo"
  
    var ss = SpreadsheetApp.getActive();
    var allsheets = ss.getSheets();
  
    var exclude_sheets = ["Metrics","Moderators"];
    var inhouse_mods = ["Amanda Bustard (GQ)", "Nivetha Velupur (GQ)","Kate Polek (GQ)"]
    var external_mods = ["Lorena Araujo"]
  
    for(var s in allsheets){
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRange)
      var values = range.getValues()
  
      if (sheet.getName().startsWith(month)) { // If sheet matches month
        if (exclude_sheets.includes(sheet.getName())) continue; // Skip sheets in exclude list
  
          values.forEach(function(row) {
            row.forEach(function(col) {
              if (inhouse_mods.includes(col) && row[0,5].startsWith(meetingType)) {
                inCount +=1}
              else if (external_mods.includes(col) && row[0,5].startsWith(meetingType)) {
                exCount += 1}
              else if (inhouse_mods.includes(col) && !row[0,5].startsWith(association) && !row[0,5].startsWith(ab_condo)) {
                inCountTwo += 1}
              else if (external_mods.includes(col) && !row[0,5].startsWith(association) && !row[0,5].startsWith(ab_condo)) {
                exCountTwo += 1}
          });
        });
      }
    } // End of Loop
  
    if (modType.startsWith("Internal",28)) {
      return inCount}
    else if (modType.startsWith("External",28)) {
      return exCount}
    else if (modType.startsWith("Internal",31)) {
      return inCount}
    else if (modType.startsWith("External",31)) {
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