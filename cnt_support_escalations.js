function COUNT_SUPPORT_ESCALATIONS(l1RangeA,l2RangeA,month,supportType) {

  var l1Count = 0
  var l2Count = 0

  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();

  for(var s in allsheets){
    var sheet = allsheets[s];
    var l1RangeB = sheet.getRange(l1RangeA)
    var l2RangeB = sheet.getRange(l2RangeA)

    var l1Values = l1RangeB.getValues()
    var l2Values = l2RangeB.getValues()

    if (sheet.getName().startsWith(month)) { 

      l1Values.forEach(function(row) {
        row.forEach(function(col) {
          if (col === true) {l1Count +=1}
          else {}
        });
      });

      l2Values.forEach(function(row) {
        row.forEach(function(col) {
          if (col === true) {l2Count +=1}
          else {}
        });
      });
    
    } 
  // End of main loop
  } if (supportType === "# of Escalations to L1") {return l1Count}
    else if (supportType === "# of Escalations to L2") {return l2Count}
    else {}
}