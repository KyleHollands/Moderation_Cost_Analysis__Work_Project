function LOOP_THROUGH_SHEETS(rangeOne,rangeTwo,countOne,countTwo) {

    var count = 0
    var cOne = 0
    var cTwo = 0
  
    // Sheets to exclude (array)
    var exclude = ["Meeting Volume","Moderators"];
  
    var ss = SpreadsheetApp.getActive(); // Acquire active sheet
    var allsheets = ss.getSheets();
  
    for(var s in allsheets){
      var sheet = allsheets[s];
  
      // Stop iteration execution if the condition is met.
      if(exclude.indexOf(sheet.getName()) == 1) continue;
        
        sheet.getRange(rangeOne)
          .getValues()
          .reduce(function (a, b) {
            rOne = a.concat(b)
            return rOne;
          })
  
        sheet.getRange(rangeTwo)
          .getValues()
          .reduce(function (a, b) {
            rTwo = a.concat(b)
            return rTwo;
          })
  
        rOne.forEach(function (v) {
          if (v === countOne) {
            cOne += 1;
          }
        })
        rTwo.forEach(function (x) {
          if (x === countTwo) {  
            cTwo += 1;
          } 
        })
          //count += 1;  
            
          
    } // End of loop
    return cOne + cTwo;
  
  } // End of function
  
  
  
  
  
  
  
  