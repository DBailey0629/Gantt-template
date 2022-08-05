/**
 * Listens for changes to start or end dates, duration,
 * or edits to the gantt chart itself
 */
function onEdit(e) {
  let sheet = e.source.getActiveSheet();
  if (sheet.getRange('A1').getValue() != "Gantt Chart") { return; }
  if (sheet.getRange('E8').getValue() == false) { return; }

  let rowEdit = e.range.getRow();
  let colEdit = e.range.getColumn();
  if (e.range.getA1Notation() == 'D7' || e.range.getA1Notation() == 'F7') { setupNewStart(); return; }
  if (colEdit < 5 || rowEdit < 13) { return; }

  if (colEdit == 5 || colEdit == 6) { datesChanged(e); return;}
  if (colEdit == 7) { durationChanged(e); return;}
  if (colEdit >= 8 ) { chartChanged(e); return;}
//       sheet.getRange('F6').setValue(colEdit);
}


/**
 * Changes the colored cells in chart based on new start
 * and end dates.
 */
function datesChanged(e) {
//  var ui = SpreadsheetApp.getUi(); // Same variations.
//  ui.alert(JSON.stringify(e));
  let sheet = e.source.getActiveSheet();
  let rowEdit = e.range.getRow();
  let colorPicker = sheet.getRange('C8').getBackground();
  let maxCols = sheet.getLastColumn();
  let dateRange = sheet.getRange(10, 8, 1, maxCols-8).getValues();
  let dateBegin = new Date(sheet.getRange(rowEdit, 5).getValue());
  let dateEnd = new Date(sheet.getRange(rowEdit, 6).getValue());
 //   sheet.getRange('F6').setValue(dateBegin+" "+dateEnd);

//Check if dates are valid
  if (isNaN(dateBegin) || isNaN(dateEnd)) {
    sheet.getRange(rowEdit, 7).setValue(null);
    sheet.getRange(rowEdit, 8, 1, maxCols).setBackground(null);
    return;
  }


  if (dateEnd < dateBegin) {return;}
  
  let colStart = 0;
  let colEnd = 0;
  dateRange[1] = [];

  dateRange[1] = dateRange[0].map(n=>n.getTime());
  colStart = dateRange[1].indexOf(dateBegin.getTime());
  colEnd = dateRange[1].indexOf(dateEnd.getTime());
  
  if (colStart == -1 || colEnd == -1) {return;}
  
  sheet.getRange(rowEdit, 8, 1, colStart).setBackground(null);
  sheet.getRange(rowEdit, colStart+8, 1, colEnd-colStart+1).setBackground(colorPicker);
  sheet.getRange(rowEdit, 8+colEnd+1, 1, dateRange[0].length-colEnd+1).setBackground(null);

  if(sheet.getRange(rowEdit, 7).getValue == null) {
    sheet.getRange(rowEdit, 7).setFormula("=max(days(R[0]C[-1],R[0]C[-2])+1,1)");
  }
}


/**
 * Listens for changes to duration, recalculates end date,
 * then calls datesChanged to change colored cells
 */
function durationChanged(e) {
  let sheet = e.source.getActiveSheet();
  let rowEdit = e.range.getRow();
  let dateBegin = new Date(sheet.getRange(rowEdit, 5).getValue());
  let dateEnd = new Date(sheet.getRange(rowEdit, 6).getValue());
  let duration = sheet.getRange(rowEdit, 7).getValue();
  
//Check if dates are valid
  if ((duration == null || duration == 0) && (isNaN(dateBegin) || isNaN(dateEnd))) {
    sheet.getRange(rowEdit, 8, 1, maxCols).setBackground(null);
    return;
  }

  dateEnd = new Date(dateBegin.getFullYear(), dateBegin.getMonth(), dateBegin.getDate() + duration - 1);
  sheet.getRange(rowEdit, 6).setValue(dateEnd);

  datesChanged(e);

  sheet.getRange(rowEdit, 7).setFormula("=max(days(R[0]C[-1],R[0]C[-2])+1,1)");
}


/**
 * If the cells in the chart are changed, match start and end dates
 * with colored cells.
 */
function chartChanged(e) {
  let colorPicker = e.source.getRange('C8').getBackground();
  let sheet = e.source.getActiveSheet();
  let range = e.range.getA1Notation();
  let rowEdit = e.range.getRow()
  let maxCols = sheet.getLastColumn();
  let dateRange = sheet.getRange(10, 8, 1, maxCols-8).getValues();
  let pointer = 0;
  let startDate = null;
  let endDate = null;

  //find start date (first cell to be color, but not red)
  while (pointer<dateRange[0].length && sheet.getRange(rowEdit, (8+pointer),1,1).getBackground() == (null || "#ffffff")) {
    pointer++;
  }    
  startDate = dateRange[0][pointer];
 
  //find end date (first cell to not be color and not red - 1
  while (pointer<dateRange[0].length && sheet.getRange(rowEdit, (8+pointer),1,1).getBackground() !== "#ffffff") {
    pointer++;
  }
  pointer--;
  if (sheet.getRange(rowEdit, (8+pointer),1,1).getBackground()=="#ff0000") {pointer--;}
  endDate = dateRange[0][pointer];

//     sheet.getRange('F8').setValue(endDate);
 
  sheet.getRange(rowEdit, 5).setValue(startDate);
  sheet.getRange(rowEdit, 6).setValue(endDate);
  sheet.getRange(rowEdit, 8+pointer+1, 1, dateRange[0].length-8-pointer).setBackground(null);
  
//Make sure that conditional formatting doesn't get screwed up  
  sheet.clearConditionalFormatRules();
  range = sheet.getRange(13,8,sheet.getLastRow(),maxCols);
  var rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=H$11=1")
  .setBackground("#FF0000")
  .setRanges([range])
  .build();
  sheet.setConditionalFormatRules([rule]);
}


/**
 * Changes the start date of the Gantt Chart
 * 
 */
function setupNewStart() {
  if (SpreadsheetApp.getActiveSheet().getRange('A1').getValue() != "Gantt Chart") { return }
  
  var startDate = new Date(SpreadsheetApp.getActiveSheet().getRange('D7').getValue());
  var endDate = new Date(SpreadsheetApp.getActiveSheet().getRange('F7').getValue());
    if (startDate >= endDate) { return; }
  var dateDiff =  Math.round(( endDate-startDate ) / 86400000) + 1.0; 
  var maxCols = SpreadsheetApp.getActiveSheet().getMaxColumns();
  var maxRows = SpreadsheetApp.getActiveSheet().getMaxRows();
  
  if (dateDiff+7 > maxCols) { SpreadsheetApp.getActiveSheet().insertColumnsAfter(maxCols, (dateDiff+7-maxCols));}
  if (dateDiff+7 < maxCols) { 
      SpreadsheetApp.getActiveSheet().getRange(6,8,2, maxCols-7).breakApart();
      SpreadsheetApp.getActiveSheet().deleteColumns((7+dateDiff+1), (maxCols-(7+dateDiff)));
  }
  maxCols = SpreadsheetApp.getActiveSheet().getMaxColumns();
  
  
//  Logger.log(dateDiff);
  
  var dateArray = [];
  var formulasDate = [];
  var formulasDay = [];
  var formulasHoliday = [];
  for (var d = 0; d < dateDiff; d++) { 
    dateArray.push(new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()+d));
    formulasDate.push("=DAY(R[-2]C[0])");
    formulasDay.push("=SWITCH(WEEKDAY(R[2]C[0]),1,\"S\",2,\"M\",3,\"T\",4,\"W\",5,\"R\",6,\"F\",7,\"S\")");
    formulasHoliday.push("=sum(arrayformula((R[-1]C[0]=Holidays!$B$2:$L)*1))");
  }

//  Logger.log(dateArray.length);
  SpreadsheetApp.getActiveSheet().getRange(10, 8, 1, dateDiff).setValues([dateArray]);
  SpreadsheetApp.getActiveSheet().getRange(12, 8, 1, dateDiff).setFormulas([formulasDate]);
  SpreadsheetApp.getActiveSheet().getRange(8, 8, 1, dateDiff).setFormulas([formulasDay]);
  SpreadsheetApp.getActiveSheet().getRange(11, 8, 1, dateDiff).setFormulas([formulasHoliday]);

  //Set up week numbers and merge correctly
  SpreadsheetApp.getActiveSheet().getRange(7,8,1, dateDiff+1).breakApart();
  var pointer = 8;
  var startDateDay = startDate.getDay();
  SpreadsheetApp.getActiveSheet().getRange(7, pointer, 1, 7-startDateDay).setValue("Week 1").merge();
  SpreadsheetApp.getActiveSheet().getRange(7, pointer, 2, 7-startDateDay)
      .setBackgroundRGB(Math.round(Math.random()*255),Math.round(Math.random()*255),Math.round(Math.random()*255));
  pointer = pointer+(7-startDateDay);
  
  var week = 2;
  while (pointer < (7+dateDiff)) {
    SpreadsheetApp.getActiveSheet().getRange(7, pointer, 1, 7).merge().setValue("Week "+week);
    SpreadsheetApp.getActiveSheet().getRange(7, pointer, 2, 7)
      .setBackgroundRGB(Math.round(Math.random()*255),Math.round(Math.random()*255),Math.round(Math.random()*255));
    pointer = pointer+7;
    week++;
  }
  
  //Clean up extra days at end
  maxCols = SpreadsheetApp.getActiveSheet().getMaxColumns();
  if (dateDiff+7 < maxCols) { SpreadsheetApp.getActiveSheet().deleteColumns((7+dateDiff+1), (maxCols-(7+dateDiff)));}
  maxCols = SpreadsheetApp.getActiveSheet().getMaxColumns();

  //Set up months and merge correctly
  SpreadsheetApp.getActiveSheet().getRange(6,8,1, maxCols-7).breakApart();
  pointer = 8;
  var startDay = startDate.getDate();
  var daysInMonth = new Date();
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  daysInMonth.setFullYear(startDate.getYear(), (startDate.getMonth()+1), 0);
  SpreadsheetApp.getActiveSheet().getRange(6, pointer, 1, daysInMonth.getDate()-startDay+1).merge().setValue(months[startDate.getMonth()]);
  pointer = pointer+(daysInMonth.getDate()-startDay)+1;
  var t = 1;
  while (pointer < (7+dateDiff+1)) {
    daysInMonth.setFullYear(startDate.getYear(), (startDate.getMonth()+t+1), 0);
    Logger.log(daysInMonth.getDate());
    SpreadsheetApp.getActiveSheet().getRange(6, pointer, 1, daysInMonth.getDate()).merge().setValue(months[daysInMonth.getMonth()])
      .setBackgroundRGB(Math.round(Math.random()*255),Math.round(Math.random()*255),Math.round(Math.random()*255));
    pointer = pointer+daysInMonth.getDate();
    t++;
  }
  
  //Clean up extra days at end
  maxCols = SpreadsheetApp.getActiveSheet().getMaxColumns();
  if (dateDiff+7 < maxCols) { SpreadsheetApp.getActiveSheet().deleteColumns((7+dateDiff+1), (maxCols-(7+dateDiff)));}
  maxCols = SpreadsheetApp.getActiveSheet().getMaxColumns();

  //set borders
  SpreadsheetApp.getActiveSheet().getRange(13, 8, maxRows-12, dateDiff).setBorder(null, null, true, true,true,true,"#d9d9d9",SpreadsheetApp.BorderStyle.DOTTED);  
  
  //Make sure that conditional formatting doesn't get screwed up  
  let sheet = SpreadsheetApp.getActiveSheet();
  sheet.clearConditionalFormatRules();
  let range = sheet.getRange(13,8,sheet.getLastRow(),maxCols);
  let rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=H$11>0")
  .setBackground("#FF0000")
  .setRanges([range])
  .build();
  sheet.setConditionalFormatRules([rule]);

}
