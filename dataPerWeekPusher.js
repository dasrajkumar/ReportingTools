//@author: Adonis Settouf
//@mail: adonis.settouf@gmail.com


var sheetId = "1k8_1boJa8lMiKmYaJNFE0iN5xZn7XAFDpbew-C1a0ns"
//Spreadsheet result
var spR = SpreadsheetApp.openById(sheetId).getActiveSheet();
//Spreadsheet source
var spSo = SpreadsheetApp.getActiveSpreadsheet();
//in which row we should write
var rowTo = 1;
//date to check entries ( by the way, check only on first and last sheet make sense)
var arrDateMonth = utilityDateMonth_();
//placeholder to retrieve correct sheet in the range 
var firstDayWeekNo;
var lastDayWeekNo;
var firstWeekNo;
var lastWeekNo;
var numberColumns = 15;

//main function executing the the big loop
function run(){
  spR = createSheet2_(spR, sheetId);
  init2_(spR,1);
  Logger.log("row: " + rowTo);
  var currWeek = returnWeekNumber_(new Date());
  Logger.log("currWeek: " + currWeek);
  routineForWriting_(currWeek, true);
  Logger.log("processing done");
  spR.getRange(rowTo,1).setValue("");
}

function init2_(sps, row){
  row = row ? row: 1;
  var col = 1;
  //15 columns, discard labels after
  writeDataInCell_(sps, "MailAddress",row, col++);
  writeDataInCell_(sps, "Date",row, col++);
  writeDataInCell_(sps, "Week",row,col++);
  writeDataInCell_(sps, "Message subject",row,col++);
  writeDataInCell_(sps, "From",row,col++);
  writeDataInCell_(sps, "To",row,col++);
  writeDataInCell_(sps, "Cc",row,col++);
  writeDataInCell_(sps, "Type",row,col++);
  writeDataInCell_(sps, "Type2",row,col++);
  writeDataInCell_(sps, "Unread",row,col++);
  writeDataInCell_(sps, "Time since Unread",row,col++);
  writeDataInCell_(sps, "Number of mails per thread",row,col++);
  writeDataInCell_(sps, "Day of the week",row,col++);
  writeDataInCell_(sps, "Hour of the day",row,col++);
  writeDataInCell_(sps, "SLA",row,col++);
  writeDataInCell_(sps, "Time to answer",row,col++);
  while (sps.getRange(rowTo, 1).getValue() != ""){ //no overwriting of existing values
    rowTo++;
  }
  boldingFirstRow_(sps, col + 1, row);
}
//create sheet on diff spreadsheet
function createSheet2_(sps, id){
  var curr = new Date();
  var currMonth = (returnWeekNumber_(curr) - 1).toString();
  Logger.log("Week: " + currMonth);
  sps.getName().indexOf("Week") != -1 ? "": sps.setName("Week "+ currMonth);
  if (sps.getRange(2,1).getValue() != ""){
    var currSheet = SpreadsheetApp.openById(id);
    var sp = currSheet.getSheetByName("Week "+ currMonth) ? 
     currSheet.getSheetByName("Week "+ currMonth): currSheet.insertSheet( "Week "+ currMonth);
    return sp;
  }
  return sps;
}

//master execution
function routineForWriting_(i, isFirstOne){
  var sheetNameToGet = "Datas for week: " + i.toString()
  Logger.log("name: " + sheetNameToGet);
  var sheet = spSo.getSheetByName(sheetNameToGet) ? spSo.getSheetByName(sheetNameToGet): spSo.getSheetByName(sheetNameToGet + " (Siesta mode)");
  Logger.log("name: " + sheet);
  if (!sheet){
    return;
  }
  
  var sheetVal = checkValuesMonth_(i,sheet.getSheetValues(2, 1, sheet.getLastRow(), sheet.getLastColumn()));
  writeToSheet_(sheetVal);
}



//create an array with current month and last month date in string format
function utilityDateMonth_(){
  var curr = new Date(); 
  var y = curr.getFullYear();
  var m = curr.getMonth();
  var firstDay = new Date(y, m, 1);
  var dateLastMonth = new Date(y, m + 1, 0);
  var firstDayMonth = Utilities.formatDate(firstDay, "CET", "yyyy/MM/dd");
  var lastDayMonth = Utilities.formatDate(dateLastMonth,"CET", "yyyy/MM/dd");
  lastDayWeekNo = mailWeek_(dateLastMonth);
  firstDayWeekNo = mailWeek_(firstDay);
  var result = [firstDayMonth,lastDayMonth];
  return result;
}

//date is as a string here
//return if the line should be discarded or not
function lineToDiscard_(data, currWeek){
  Logger.log("currWeek: " + currWeek);
  Logger.log("week ret: " + data[1]);
  if(currWeek.toString().indexOf(data[1]) != -1){
    Logger.log("true");
    return true;
  } else {
    return false;
  }
}

//datas is a two dimensional array
function checkValuesMonth_(currWeek,datas){
  var i = 0;
  var res = [];
  var currMonth = new Date();
  currMonth = currMonth.getMonth() + 1;
  for(i; i<datas.length; i++){ 
    lineToDiscard_(datas[i], currWeek - 1) ? res.push(datas[i]): "";
  }
  Logger.log("res: " + res);
  return res;
}

//write datas in sheet
function writeToSheet_(datas){
  var j;
  var k;
  j = 0;
  rowTo = rowTo <= 1 ? 2: rowTo;
  for(j; j < datas.length; j++){
    k = 0;
    for(k;k <numberColumns; k++){
      spR.getRange(rowTo, 1).setValue(alias);
      spR.getRange(rowTo, k + 2).setValue(datas[j][k]);
    }
    rowTo++;
  }
  rowTo--;
}

//doStuff
function retrieveValueInRange_(){
  Logger.log("in retrieveValueFunc");
  var i = firstDayWeekNo +1;
  routineForWriting_(i,true);
  i++;
  for (i; i < lastDayWeekNo + 1; i++){
    routineForWriting_(i, true);
  }
}

//calculate ISO week number (thanks to the author of this awesome piece of code!)
function returnWeekNumber_(date){
  var target  = date;  
  
  // ISO week date weeks start on monday  
  // so correct the day number  
  var dayNr   = (date.getDay() + 6) % 7;  
  
  // ISO 8601 states that week 1 is the week  
  // with the first thursday of that year.  
  // Set the target date to the thursday in the target week  
  target.setDate(target.getDate() - dayNr + 3);  
  
  // Store the millisecond value of the target date  
  var firstThursday = target.valueOf();  
  
  // Set the target to the first thursday of the year  
  // First set the target to january first  
  target.setMonth(0, 1);  
  // Not a thursday? Correct the date to the next thursday  
  if (target.getDay() != 4) {  
    target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);  
  }  
  
  // The weeknumber is the number of weeks between the   
  // first thursday of the year and the thursday in the target week  
  var weekNo = 1 + Math.ceil((firstThursday - target) / 604800000);
  return weekNo; 
}