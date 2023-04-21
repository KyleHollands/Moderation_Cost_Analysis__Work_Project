function CALCULATE_MOD_COST(dataRange, month, modType) {

  var internalModCost = 0
  var externalModCost = 0
  var externalModCostFifteenPerc = 0
  var externalModCostThirtyPerc = 0
  var allModCost = 0
  var exCount = 0

  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();

  //var exclude_sheets = ["Moderators","MOps OKRs","Meeting Volume"];
  var internal_mods = ["Amanda Bustard (GQ)", "Nivetha Velupur (GQ)","Kate Polek (GQ)","Rebecca Trotter (GQ)"]

  var external_mods = ["Stacy Costa",	"Sachi Lovatt",	"Chantel Espinola",	"Amy Stephenson",	"Lorena Araujo",	"Joanne Janzen",	"Patricia Walker",	"Rebecca Blankfort",	"Becah Machado-Burton",	"Susan Hwang",	"Sandra Tudge",	"Olivia Dutka",	"Caleb Hayward",	"Cody Sharpe",	"Steph Jacobs",	"Amrinder Singh",	"Wendy Dennis",	"Bhavani Kidambi",	"April Scarff",	"Arshpreet Kaur",	"John Harris",	"Gabriela Jungblut","Eric Aspila","Kathleen Speckert","Liona Davies","Eric Szyiko","Andrew Natale","Sheila McFadyen","Leanne Unruh","Sharry Ash","Jaspreet Singh","Eric Gloutnez","Justin Trent","Ren Haskett","Andreane Dicaire","Vidhi Abbi","Fatima Mehrzad","Elisa Brosseau"]
  
  var all_mods = ["Stacy Costa",	"Sachi Lovatt",	"Chantel Espinola",	"Amy Stephenson",	"Lorena Araujo",	"Joanne Janzen",	"Patricia Walker",	"Rebecca Blankfort",	"Becah Machado-Burton",	"Susan Hwang",	"Sandra Tudge",	"Olivia Dutka",	"Caleb Hayward",	"Cody Sharpe",	"Steph Jacobs",	"Amrinder Singh",	"Wendy Dennis",	"Bhavani Kidambi",	"April Scarff",	"Arshpreet Kaur",	"John Harris",	"Gabriela Jungblut","Eric Aspila","Kathleen Speckert","Liona Davies","Eric Szyiko","Andrew Natale","Sheila McFadyen","Leanne Unruh","Sharry Ash","Jaspreet Singh","Eric Gloutnez","Justin Trent","Ren Haskett","Andreane Dicaire","Vidhi Abbi","Fatima Mehrzad","Elisa Brosseau","Amanda Bustard (GQ)", "Nivetha Velupur (GQ)","Kate Polek (GQ)","Rebecca Trotter (GQ)"]

  for(var s in allsheets){
    var sheet = allsheets[s]
    var range = sheet.getRange(dataRange)
    var values = range.getValues()

    // If sheet matches month
    if (sheet.getName().startsWith(month)) { 

      // Current Moderator Pricing Model
      values.forEach(function(row) {
        row.forEach(function(col) {
          if (internal_mods.includes(col) && internal_mods.includes(row[0,0])) {internalModCost += 1 * 30 * 2.5}
          else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,5].startsWith("Association")) {externalModCost += 350, exCount += 1}
          else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] <= 149) {externalModCost += 260, exCount += 1}
          else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] >= 150) {externalModCost += 300, exCount += 1}
          else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] == "") {externalModCost += 280, exCount += 1}
          else {}
        });
      });

      // Cost if all Moderators were External at Current Price Tiers
      values.forEach(function(row) {
        row.forEach(function(col) {
          if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,5].startsWith("Association")) {allModCost += 350}
          else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] <= 149) {allModCost += 260}
          else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] >= 150) {allModCost += 300}
          else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] == "") {allModCost += 280}
          else {}
        });
      });

      // Cost of External Moderators at Various Tiers
    } potentialModCostIfExternalInternal = exCount * 30 * 2.5
      externalModCostFifteenPerc = exCount * (30 * 1.15) * 2.5
      externalModCostThirtyPerc = exCount * (30 * 1.30) * 2.5
      externalModCostFortyFivePerc = exCount * (30 * 1.45) * 2.5
      externalModCostSixtyPerc = exCount * (30 * 1.60) * 2.5

  } // End of Loop

  if (modType.startsWith("Internal Mods - Approx. Cost")) {return internalModCost}
  else if (modType.startsWith("Total - Approx. Potential Cost if all External")) {return allModCost}
  else if (modType.startsWith("External Mods - Approx. Cost")) {return externalModCost}
  else if (modType.startsWith("Potential Cost if External Mods at +15%")) {return externalModCostFifteenPerc}
  else if (modType.startsWith("Potential Cost if External Mods at +30%")) {return externalModCostThirtyPerc}
  else if (modType.startsWith("Potential Cost if External Mods at +45%")) {return externalModCostFortyFivePerc}
  else if (modType.startsWith("Potential Cost if External Mods at +60%")) {return externalModCostSixtyPerc}
  else if (modType.startsWith("Potential Cost Diff")) {return potentialModCostIfExternalInternal}
  else {}
} 