/******************************
Written by Anders Ohman, 2015-16
for the Dinulescu Lab at BWH.
andy.ohman@gmail.com

Main Scripts:
1) onOpen()
   Create new file menu
2) autoParse()
   Process new tag form responses
3) updateTotal()
   Update newest WAB totals
4) moveToThorn(#)
   Move a mouse to Thorn
5) sacMouse(#)
   Sacrifice a mouse
6) hideDead() - KIND OF WORKING
   Hide dead mice from sheet
7) getInfo(#)
   Get info card on a mouse

Helper Methods:
a) inArray(#,A)
   Check if item is in an array
b) parseNumbers(#-#)
   Parses list of mouse #s into array
c) numOrList(#-#,-)
   Determines if range is a # or '#-#'
d) makeRangeList(#-#)
   Make a dash-separated # range into array
e) find(#,range)
   Find a value in a sheet range
f) newRow(cohort,list,males,females,input)
   Makes a new row in GT sheet for a mouse
g) findNum(#,sheet)
   Searches for a # and finds location
h) findEnd(cell,sheet)
   Determines appropriate EndTarget location

NONE OF THESE SCRIPTS WILL FUNCTION
UNTIL PLACED IN SHEET'S SCRIPT EDITOR!
******************************/

/* 1: MENU CREATION */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('GT Scripts')
  .addItem('Process New Tags from Form', 'autoParse')
  .addItem('Update WAB Total #s', 'updateTotal')
  .addItem('Get Mouse Info', 'getInfo')
  .addSubMenu(ui.createMenu('Mouse Management')
              .addItem('Move a Mouse to Thorn', 'moveToThorn')
              .addItem('Sacrifice a Mouse', 'sacMouse')
              .addItem('Hide dead mice', 'hideDead'))
  .addToUi();
}
/* 2: PARSE NEW TAG FORM RESPONSE INTO SHEET */
function autoParse() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var s = ss.getSheetByName("Form Responses 1");
var r = s.getRange(s.getLastRow(), 1);

if(r.getValue() == 'Timestamp') r = r.offset(1,0); // If on the title row, increment to row 2
if(!r.isBlank()) { // If on the response sheet, and row 2 isn't blank
  var row = r.getRow(); // Get current row
  var numColumns = s.getLastColumn(); // Get total number of columns in row
  var valuesArray = s.getRange(row, 1, 1, numColumns).getValues(); // Get all values from row

  if(typeof valuesArray === 'object'){ //if the array is an array, just an error check
    var genotype = valuesArray[0][1];
    var bday = valuesArray[0][2];
    var mom = valuesArray[0][3];
    var dad = valuesArray[0][4];
    var range = valuesArray[0][5];
    var males = valuesArray[0][6];
    var females = valuesArray[0][7];

    var newList = parseNumbers(range);
    var newMales = parseNumbers(males);
    var newFemales = parseNumbers(females);

    var targetSheet = ss.getSheetByName(genotype); // Determines destination tab based on cohort answer
    while(newList.length > 0) newRow(targetSheet, newList.pop(), newMales, newFemales, valuesArray);
    s.deleteRow(row);  // Purge Form Response row after completion
    Browser.msgBox('New Tag Form Auto-Processer','Added '+range+' to sheet '+targetSheet.getName()+'.',Browser.Buttons.OK); // Report result
  }
  else Browser.msgBox('New Tag Form Auto-Processer','There is an issue with the form response.',Browser.Buttons.OK);
}
  else Browser.msgBox('New Tag Form Auto-Processer','No new tags to add.',Browser.Buttons.OK);
}
/* 3: UPDATE TOTAL WAB TRENDLINE */
function updateTotal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = new Date();
  var todaysDate = d.toLocaleDateString();
  var mouseTotal = ss.getSheetByName('Coversheet').getRange(2,8).getValue();
  var target = ss.getSheetByName('Data').getRange(ss.getSheetByName('Data').getLastRow(), 2); // Goes to the last populated row
  if (todaysDate == target.getValue().toLocaleDateString()) Browser.msgBox('WAB Totals Updater','Today already processed. Please make changes manually.',Browser.Buttons.OK);
  else {
    target.offset(1,0).setValue(todaysDate);
    target.offset(1,1).setValue(mouseTotal);
    Browser.msgBox('WAB Totals Updater','On '+todaysDate+' there were '+mouseTotal+' mice in WAB.',Browser.Buttons.OK);
  }
}
/* 4: MOVE A MOUSE TO THORN */
function moveToThorn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = new Date();
  var todaysDate = d.toLocaleDateString();
  var mouseNum = Browser.inputBox('Move Mouse To Thorn','Enter mouse number',Browser.Buttons.OK_CANCEL);

  if (mouseNum != 'cancel'){
    mouseNum = parseNumbers(mouseNum);
    while (mouseNum.length > 0){
      var currentNum = mouseNum.pop();
      var numResult = findNum(currentNum, ss);
      var targetSheet = numResult[0];
      var cell = numResult[1];
      var endTarget = numResult[2];

      if (cell != null) {
        var target = ss.getSheetByName('Thorn Transfers').getRange(ss.getSheetByName('Thorn Transfers').getLastRow() + 1, 1);
        target.setValue(cell.getValue()); // Set transferred mouse #
        target.offset(0,1).setValue(endTarget.offset(0,1).getValue()); // Set the genotype information
        target.offset(0,2).setValue(todaysDate); // Set today's date as transfer date
        endTarget.setValue(''); // Remove Alive flag from origin sheet
        endTarget.offset(0,-1).setValue(endTarget.offset(0,-1).getValue().concat(', THORN '+todaysDate)); // Sets comment about Thorn with date
        targetSheet.hideRow(cell); // Hide row on origin sheet

        Browser.msgBox('Move Mouse To Thorn','Mouse '+cell.getValue()+' in group '+targetSheet.getName()+' transferred to Thorn.',Browser.Buttons.OK);
      }
      else Browser.msgBox('Move Mouse To Thorn','Mouse # not found.',Browser.Buttons.OK);
    }
  }
}
/* 5: SACRIFICE A MOUSE */
/* It would be nice to only search non-hidden rows, but this is apparently still a feature request...
http://stackoverflow.com/questions/6793805/how-to-skip-hidden-rows-while-iterating-through-google-spreadsheet-w-google-app
*/
function sacMouse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = new Date();
  var todaysDate = d.toLocaleDateString();
  var mouseNum = Browser.inputBox('Sacrifice Mouse','Enter mouse number',Browser.Buttons.OK_CANCEL);

  if(mouseNum != 'cancel'){
    mouseNum = parseNumbers(mouseNum);
    while (mouseNum.length > 0){
      var currentNum = mouseNum.pop();
      var numResult = findNum(currentNum, ss);
      var targetSheet = numResult[0];
      var cell = numResult[1];
      var endTarget = numResult[2];

      if (cell != null) {
        if (endTarget.getValue() == 'A'){
          endTarget.clearContent(); // Remove Alive flag from origin sheet
          endTarget.offset(0,-1).setValue(endTarget.offset(0,-1).getValue().concat(', SAC on '+todaysDate)); // Sets comment about sac with date
          targetSheet.hideRow(cell); // Hide row on origin sheet
          Browser.msgBox('Sacrifice Mouse','Mouse '+cell.getValue()+' in group '+targetSheet.getName()+' sacrificed.',Browser.Buttons.OK);
        }
        else Browser.msgBox('Sacrifice Mouse','Mouse already sacrificed.',Browser.Buttons.OK);
        //if row isn't already hidden
        targetSheet.hideRow(cell); // Hide row on origin sheet
      }
      else Browser.msgBox('Sacrifice Mouse','Mouse # not found.',Browser.Buttons.OK);
    }
  }
}
/* 6: HIDE DEAD MICE FROM SHEET */
/* It would be nice to only search non-hidden rows, but this is apparently still a feature request...
http://stackoverflow.com/questions/6793805/how-to-skip-hidden-rows-while-iterating-through-google-spreadsheet-w-google-app
*/
function hideDead() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getRange('A2');

  while (range.getRowIndex() <= sheet.getLastRow()){
    var endTarget = findEnd(range, sheet);
    if (typeof range.getValue() === 'number' && endTarget.isBlank()) sheet.hideRow(range);
    if (range.getRowIndex() != sheet.getLastRow()) range = range.offset(1,0);
  }
}
/* 7: GET INFO ON MOUSE BY # */
function getInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mouseNum = Browser.inputBox('Mouse Info','Enter mouse number',Browser.Buttons.OK_CANCEL);
  if(mouseNum != 'cancel'){
    var mouseCell = findNum(mouseNum, ss);
    if (mouseCell[1] != null){
      var birthdate = new Date(mouseCell[1].offset(0,2).getValue());
      Browser.msgBox('Mouse Info: # '+mouseCell[1].getValue()+' '+mouseCell[2].getValue(),
        'Sex: '+mouseCell[1].offset(0,1).getValue()+
          '\\n Born: '+birthdate.toLocaleDateString()+
            '\\n GT: '+mouseCell[2].offset(0,1).getValue()+
              //'\\n Cohort: '+mouseCell[0].getName()+
              '\\n Parents: '+mouseCell[2].offset(0,-3).getValue()+' x '+mouseCell[2].offset(0,-2).getValue()+
                '\\n Notes: '+mouseCell[2].offset(0,-1).getValue()
                ,Browser.Buttons.OK);
    }
    else Browser.msgBox('Get Mouse Info','Mouse # not found.',Browser.Buttons.OK);
  }
}

/* HELPER METHODS
a) inArray(#,A)
   Check if item is in an array
b) parseNumbers(#-#)
   Parses list of mouse #s into array
c) numOrList(#-#,-)
   Determines if range is a # or '#-#'
d) makeRangeList(#-#)
   Make a dash-separated # range into array
e) find(#,range)
   Find a value in a sheet range
f) newRow(cohort,list,males,females,input)
   Makes a new row in GT sheet for a mouse
g) findNum(#,sheet)
   Searches for a # and finds location
h) findEnd(cell,sheet)
   Determines appropriate EndTarget location
*/
/* a: Check if an item is in an array */
function inArray(needle, haystack) {
  if(typeof haystack === 'number' && needle == haystack) return true; //if search term is the only element
  else {
    if (typeof haystack === 'object'){ //ensures haystack is an array
      var length = haystack.length;
      for(var i = 0; i < length; i++) {
        if(haystack[i] == needle) return true; }
    }
    return false; }
}
/* b: Parses list of mouse #s into arrays (including commas and dashes) */
function parseNumbers(textInput){
  if (typeof textInput === 'number') {
    var number = [];
    number.push(textInput);
    return number;
  }
  else {
    var numberArrayBefore = textInput.split(',');
    var numberArrayAfter = [];
    numberArrayBefore.reverse();
    while (numberArrayBefore.length > 0){
      var item = numberArrayBefore.pop();

      if(item.search('-')==-1) item = parseInt(item);
      item = numOrList(item, '-');
      while (item.length > 0) numberArrayAfter.push(item.pop());
    }
    numberArrayAfter.reverse();
    return numberArrayAfter;
  }
}
/* c: Figures out if new value is a number or a string, processes accordingly. */
function numOrList(value, separator) {
  var valueArray = [];
  if (typeof value === 'undefined' || typeof value === 'object') return 0; //if it's nothing, return 0
  else if(typeof value === 'number') {
    valueArray = [value];
    return valueArray; //if a number, return the number as single array element
  }
  else if(typeof value === 'string' && value.search(separator)>-1){
    if (separator == '-') return makeRangeList(value, separator); //for range of tags
    else if (separator == ',') { //for Sac/Thorn multi-entry
      valueArray = value.split(separator);
      valueArray.reverse();
      return valueArray; //otherwise make array of #s
    }
    else Browser.msgBox("Please check the appropriate separator symbol for your numbers.");
  }
}
/* d: Feed me a dash-separated range to get an array list of those #s */
function makeRangeList(numRange, separator){
  var split = numRange.split(separator); // split[0] is first number, split[1] is last
  var list = []; // This array will contain the numbers from (and including) first to last in the range given
  if(split[0]<split[1]){ // Ensures the first # comes before the last, just in case
    for(var i=split[0]; i<=split[1]; i++){ list.push(i); }
    list.reverse(); // Reverses element order to allow popping sequentially
    return list;
  }
  else Browser.msgBox('There is something wrong with your number range.');
}
/* e: Finds a value within a given range. */
function find(value, range) {
  var data = range.getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] == value) {
        // For handling really old mice with possible # repeats:
        // var today = new Date();
        // today = today.toLocaleDateString();
        // var tooOld = new Date();
        //if (range.getCell(i+1,j+2) < 18 months ago...  {
        return range.getCell(i + 1, j + 1); }// Returns a range pointing to the first cell containing the value
    }
  }
  return null; // or null if not found
}
/* f: Adds a new row to a GT sheet table when fed a target, mouse, male and female lists, and the value array */
function newRow(targetSheet, newList, newMales, newFemales, valuesArray){
  var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1); // Goes to one row below the last populated row

  var endTarget = findEnd(target, targetSheet); //goes to Alive column
  endTarget = endTarget.offset(0,-3); //backs up to Mom column

  target.setValue(newList); // Set mouse #
  if(inArray(newList,newMales)) target.offset(0,1).setValue("M"); // Sets Male gender
  else if(inArray(newList,newFemales)) target.offset(0,1).setValue("F"); // Sets Female gender
  else target.offset(0,1).setValue("?"); // If neither gender, sets a ? for gender
  target.offset(0,2).setValue(valuesArray[0][2]); // Sets birthdate
  endTarget.setValue(valuesArray[0][3]); // Sets mom #
  endTarget.offset(0,1).setValue(valuesArray[0][4]); // Sets dad #
  endTarget.offset(0,3).setValue("A"); // Sets Alive flag
  var rowNum = target.getRow(); // Sets blank genotype template
  if(targetSheet.getName().search('Myc')>-1) endTarget.offset(0,4).setValue("=(IF(ISTEXT(D"+rowNum+"),D"+rowNum+"&\"; \",\"\")&$E$1&\"(\"&E"+rowNum+"&\"); \"&IF(ISTEXT(G"+rowNum+"),$F$1&\"(\"&F"+rowNum+"&G"+rowNum+"&\"); \",$F$1&\"(\"&F"+rowNum+"&\"); \")&$H$1&\"(\"&H"+rowNum+"&\"); \"&IF(ISTEXT(I"+rowNum+"),I"+rowNum+"&\"; \",\"\")&IF(ISTEXT(J"+rowNum+"),J"+rowNum+"&\"\",\"\"))");
  else if(targetSheet.getName().search('KPTEN')>-1) {
    if(targetSheet.getName().search('RCTR RAPT')>-1) endTarget.offset(0,4).setValue("=(D$1&\"(\"&D"+rowNum+"&\"); \"&E$1&\"(\"&E"+rowNum+"&\"); \"&F$1&\"(\"&F"+rowNum+"&\"); \"&G$1&\"(\"&G"+rowNum+"&\"); \"&IF(ISTEXT(H"+rowNum+"),H"+rowNum+",))");
    else endTarget.offset(0,4).setValue("=(D$1&\"(\"&D"+rowNum+"&\"); \"&E$1&\"(\"&E"+rowNum+"&\"); \"&F"+rowNum+")");
    }
  else endTarget.offset(0,4).setValue("=($D$1&\"(\"&D"+rowNum+"&\"); \"&IF(ISTEXT(F"+rowNum+"),$E$1&\"(\"&E"+rowNum+"&F"+rowNum+"&\"); \",$E$1&\"(\"&E"+rowNum+"&\"); \")&$G$1&\"(\"&G"+rowNum+"&\"); \"&IF(ISTEXT(H"+rowNum+"),H"+rowNum+"&\"; \",\"\")&IF(ISTEXT(I"+rowNum+"),I"+rowNum+"&\"; \",\"\")&IF(ISTEXT(J"+rowNum+"),J"+rowNum+",\"\"))");
}
/* g: Search for a mouse # and return the location info  */
function findNum(mouseNum, ss){
  var cohorts = [
    "KPTEN luc",
    "KPTEN RCTR RAPT LUC",
    "Br1 p53 PTEN PaxTet",
    "Br1 p53 Myc PaxTet",
    "Br2 p53 Myc PaxTet",
    "Br2 p53 PTEN PaxTet"];

    do { // Search spreadsheets for mouse # until found
      var targetSheet = ss.getSheetByName(cohorts.pop());
      var targetColumn = targetSheet.getRange("A:A");
      var cell = find(mouseNum, targetColumn);
    }
    while (cell == null && cohorts.length > 0);

    if (cell != null) { //if the # was found
      var endTarget = findEnd(cell, targetSheet);
      return [targetSheet, cell, endTarget];
    }
    else return [targetSheet, cell, cell];
}
/* h: Determine the EndTarget location */
function findEnd(cell, targetSheet){
  var endTarget = cell.offset(0,13); // If Myc or KPTEN, have to adjust
  if(targetSheet.getName().search('KPTEN')>-1) {
    endTarget = endTarget.offset(0,-4); // Regular KPTEN
    if(targetSheet.getName().search('RCTR RAPT')>-1) endTarget = endTarget.offset(0,2); // R/R
  }
  else if(targetSheet.getName().search('Myc')>-1) endTarget = endTarget.offset(0,1); // Myc
  return endTarget;
}


function timeDifference(current, previous) {

    var msPerMinute = 60 * 1000;
    var msPerHour = msPerMinute * 60;
    var msPerDay = msPerHour * 24;
    var msPerMonth = msPerDay * 30;
    var msPerYear = msPerDay * 365;

    var elapsed = current - previous;

    if (elapsed < msPerMinute) {
         return Math.round(elapsed/1000) + ' seconds ago';
    }

    else if (elapsed < msPerHour) {
         return Math.round(elapsed/msPerMinute) + ' minutes ago';
    }

    else if (elapsed < msPerDay ) {
         return Math.round(elapsed/msPerHour ) + ' hours ago';
    }

    else if (elapsed < msPerMonth) {
        return 'approximately ' + Math.round(elapsed/msPerDay) + ' days ago';
    }

    else if (elapsed < msPerYear) {
        return 'approximately ' + Math.round(elapsed/msPerMonth) + ' months ago';
    }

    else {
        return 'approximately ' + Math.round(elapsed/msPerYear ) + ' years ago';
    }
}
