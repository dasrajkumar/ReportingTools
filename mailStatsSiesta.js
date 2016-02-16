//@author: Adonis Settouf
//@mail: adonis.settouf@gmail.com

//Spreadsheet to write the numbers
var spS = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//in which row we should write
var rowToWrite = 2;
//mail of the mailbox
var alias = Session.getActiveUser().getEmail();
//Labels
//var labels = GmailApp.getUserLabels();
//working hours, see the 9 and 18
var colLen;
var startWorkHours = 9;
var endWorkHours = 18;
//threshold for a sla to be ok (to be modified for use)
var slaThreshold = 4;
//date previous week and this week (7 days)
var arrDate = utilityDate_();
//Siesta time
var siestaTimeStart = 13;
var siestaTimeEnd = 14;

//main function executing the the big loop
/*function mainSiesta(){
  spS = createSpreadsheet_(spS);
  rowToWrite = init_(spS);
  var cache = CacheService.getScriptCache();
  var cachedCount = cache.get("lastThread");
  var rowCount = cache.get("rowToWrite");
  var i = cachedCount != null ? parseInt(cachedCount) :0;
  rowToWrite = rowCount != null ? parseInt(rowCount) : rowToWrite;
  var query = "NOT in:draft NOT in:chats NOT in:sent after:" + arrDate[0] + " before:" + arrDate[1];
  Logger.log(query);
  var threads = GmailApp.search(query);
  for(i;i < threads.length; i++){
    //routine_(threads[i],threads[i].isInTrash() ? true:false);  lolz
    keepInCache_("rowToWrite", rowToWrite);
    routine_(threads[i]);
    keepInCache_("lastThread", i.toString());
  }
}*/

//keep counter in cache for next launch
function keepInCache_(name, counter){
  var cache = CacheService.getScriptCache();
  cache.put(name, counter, 3600);
}

//utility to create new sheet each time
function createSpreadsheet_(sps){
  var curr = new Date();
  var mailWeek = (mailWeek_(curr)).toString();
  var shName = "Datas for week: "+ mailWeek + " (Siesta mode)";
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName().indexOf("Datas") != -1 ? "": SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setName(shName);
  if (sps.getRange(1,1).getValue() != ""){
    var sp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName) ? 
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName): SpreadsheetApp.getActiveSpreadsheet().insertSheet(shName);
    return sp;
  }
  return SpreadsheetApp.getActiveSheet();
}

///get things started
function init_(sps, row){
  var rowToWrite = 1;
  row = row ? row: 1;
  var col = 1;
  var i = 0;
  writeDataInCell_(sps, "Date",row, col++);
  writeDataInCell_(sps, "Week",row,col++);
  writeDataInCell_(sps, "Message subject",row,col++);
  //writeDataInCell_(sps, "Labels",row, col++);
  writeDataInCell_(sps, "From",row,col++);
  writeDataInCell_(sps, "To",row,col++);
  writeDataInCell_(sps, "Cc",row,col++);
  //writeDataInCell_(spS, "Label",row,col++);
  writeDataInCell_(sps, "Type",row,col++);
  writeDataInCell_(sps, "Type2",row,col++);
  writeDataInCell_(sps, "Unread",row,col++);
  writeDataInCell_(sps, "Time since Unread",row,col++);
  writeDataInCell_(sps, "Number of mails per thread",row,col++);
  writeDataInCell_(sps, "Day of the week",row,col++);
  writeDataInCell_(sps, "Hour of the day",row,col++);
  writeDataInCell_(sps, "SLA",row,col++);
  writeDataInCell_(sps, "Time to answer",row,col++);
  colLen = col;
   for (i = 0; i < labels.length; i++){
    writeDataInCell_(spS, labels[i].getName(),row, col++);
  }
  while (sps.getRange(rowToWrite, 4).getValue() != ""){ //no overwriting of existing values
    rowToWrite++;
  }
  boldingFirstRow_(sps, col + 1, row);
  return rowToWrite;
}

//set bold on first row
function boldingFirstRow_(sps, rowSize, row){
  var i = 1;
  for(i; i < rowSize; i++){
    sps.getRange(row,i).setFontWeight("bold");
  }
}

//write to spreadsheet all infos
function routine_(thread){
  var i = 0;
  var msgs = thread.getMessages();
  var threadName = thread.getFirstMessageSubject();
  var numOfMailPerThread = thread.getMessageCount();
  var slacounter = 0;
  var col;
  var rowForSla = [];
  var isIncoming;
  var incoming;
  var arrToCc;
  var type2;
  var dateMail;
  var arrDayAndHour;
  var isUnreadTime;
  var isReadString;
  var buffRow = [];
  var areLabelsPresent = findLabel_(thread);
  for(i;i<msgs.length;i++){
    if(msgs[i].getSubject().indexOf("Attachment") > -1){
      continue;
    }
    if(msgs[i].getDate() > arrDate[0]){
      continue;
    }
    stripGetFrom_(msgs[i].getFrom()).indexOf(alias) > -1 ? rowForSla.push(rowToWrite): "";
    col = 1;
    isIncoming = checkIncoming_Mail_(msgs[i]);
    incoming = isIncoming ? "INCOMING" : "OUTGOING";
    arrToCc = checkToCc_(msgs[i]);
    type2 = msgs[i].isInTrash() ? "Junk" : checkMailType_(msgs[i], isIncoming);
    dateMail = msgs[i].getDate();
    arrDayAndHour = findDayAndHour_(dateMail);
    isUnreadTime = isUnreadTime_(msgs[i]);
    isReadString = msgs[i].isUnread() ? "Yes" : "No";
    buffRow.push(rowToWrite);
    writeDataInCell_(spS, Utilities.formatDate(dateMail, "CET", "dd/MM/yyyy"), rowToWrite, col++);
    writeDataInCell_(spS, mailWeek_(dateMail), rowToWrite, col++);
    writeDataInCell_(spS, threadName, rowToWrite, col++);
    writeDataInCell_(spS, msgs[i].getFrom(), rowToWrite, col++);
    writeDataInCell_(spS, arrToCc[0], rowToWrite, col++);
    writeDataInCell_(spS, arrToCc[1], rowToWrite, col++);
    writeDataInCell_(spS, incoming, rowToWrite, col++);
    writeDataInCell_(spS, type2, rowToWrite, col++);
    writeDataInCell_(spS, isReadString, rowToWrite, col++);
    writeDataInCell_(spS, isUnreadTime, rowToWrite, col++);
    writeDataInCell_(spS, numOfMailPerThread, rowToWrite, col++);
    writeDataInCell_(spS,arrDayAndHour[0], rowToWrite, col++);
    writeDataInCell_(spS, arrDayAndHour[1], rowToWrite, col++);
    rowToWrite++;
  }
  rowForSla.length > 0 ? col = checkSLA_(msgs, rowForSla, col): "";
  var j = 0;
  var colLab;
  for(i=0;i<msgs.length;i++){
    colLab = colLen;
     for (j = 0; j < areLabelsPresent.length; j++){
      writeDataInCell_(spS, areLabelsPresent[j], buffRow[i], colLab++);
    }
  }
}

//returns tab with day and hour from a date
function findDayAndHour_(date){
  res = [];
  res.push(Utilities.formatDate(date, "CET", "EEEE"));
  res.push(Utilities.formatDate(date, "CET", "H"));
  return res;
}

//find labels in a thread
function findLabel_(thread){
  var labs = thread.getLabels();
  var j = 0;
  var tt = [];
  for(j;j<labs.length;j++){
    tt.push(labs[j].getName());
  }
  var i = 0;
  var res = [];
  for (i;i<labels.length;i++){
    tt.indexOf(labels[i].getName()) != -1 ? res.push("Y") : res.push("N");
  }
  return res;
}

//check if mail is read
function isUnreadTime_(mail){
  if(mail.isUnread()){
    var curr = new Date();
    var timeTaken = curr.getTime() - mail.getDate().getTime();
    Logger.log(timeTaken);
    timeTaken = (timeTaken)/(3600 * 1000);
    return timeTaken;
  }
  else{
    return "";
  }
}
//return cc and to of a mail (to in first, cc in second)
function checkToCc_(message){
  var mail = message;
  if(mail.getTo() && mail.getCc()){
    return [mail.getTo(),mail.getCc()];
  } else{
    var arr = [];
    mail.getTo() ? arr.push(mail.getTo()):arr.push("No To");
    mail.getCc() ? arr.push(mail.getCc()):arr.push("No Cc");
    return arr;
  }
}

//auxiliary function to easily write in a spreadsheet  
function writeDataInCell_(sps, data, row, col){
  sps.getRange(row, col).setValue(data);
}

//check if empty, returns false if not
function emptyChecker_(row,col){
  //lot of calls to server, perhaps change stepping (more than one range) to increase perf
  if(spS.getRange(row, col).getValue() == ""){
    return true;
  } else{
    return false;
  }
}

//return type2 of mail
function checkMailType_(mail, isIncoming){
  if(isIncoming){
    if (mail.getFrom().indexOf("lexmark") != -1){
      return "Lexmark";
    } else{
      return "Customer";
    }
  } else{
    if (mail.getTo().indexOf("lexmark") != -1){
      return "Lexmark";
    } else{
      return "Customer";
    }
  }
}

//create an array with current week and last week date in string format for gmail querying
function utilityDate_(){
  var curr = new Date(); // get current date
  var dateLastWeek = new Date(curr.getTime() - 1000*60*60*24*7);
  var thisWeek = Utilities.formatDate(curr, "CET", "yyyy/MM/dd");
  var lastWeek = Utilities.formatDate(dateLastWeek,"CET", "yyyy/MM/dd");
  var result = [lastWeek,thisWeek];
  return result;
}

//calculate simple week number
/*function mailWeek_(date){
  var yearStart = new Date(date.getYear(),0,1);
  var weekNo = Math.ceil(( ( (date - yearStart) / 86400000) + 1)/7);
  return (weekNo>52) ? 52: weekNo;
} */

//calculate ISO week number (thanks to the author of this awesome piece of code!)
function mailWeek_(date){
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
  return (weekNo > 52) ? 52: weekNo; 
}

//check if Incoming or Outgoing thread
function checkIncoming_(messages){
  var firstMessage = messages[0];
  if (firstMessage.getFrom().indexOf(alias) != -1){
    return false;
  } else{
    return true;
  }
}

//check if Incoming or Outgoing for each mail
function checkIncoming_Mail_(mail){
  if (mail.getFrom().indexOf(alias) != -1){
    return false;
  } else{
    return true;
  }
}

//Strip a sender from unnecessary signs
function stripGetFrom_(sender){
  return  (sender.indexOf("<") != -1) ? sender.split("<")[1].split(">")[0]: sender;
}

//retrieve the date of a mail under string format, for example "13/01/2014"
function getDate_(mail){
  return Utilities.formatDate(mail.getDate(), "CET", "dd/MM/yyyy");
}


//reset date to start of the shift
function resetDate_(dateIncoming){
  dateIncoming.setHours(startWorkHours);
  dateIncoming.setMinutes(0);
  dateIncoming.setSeconds(0);
  return dateIncoming;
}

//function to reset date when outside of shift time
function resetDateWithShift_(dateIncoming){
  var hCondition = (dateIncoming.getHours() < startWorkHours || dateIncoming.getHours() >= endWorkHours);
  //change hour block if necessary
  if(hCondition){
    dateIncoming.getHours() >= endWorkHours ? dateIncoming.setDate(dateIncoming.getDate() +1): "";
    dateIncoming = resetDate_(dateIncoming);
  }
  //Logger.log("after hours mod day incom: " + dateIncoming);
  //change date block if necessary
  if(dateIncoming.getDay() == 6){
    dateIncoming.setDate(dateIncoming.getDate() + 2);
    dateIncoming = resetDate_(dateIncoming);
  } else if(dateIncoming.getDay() == 0){
    dateIncoming = resetDate_(dateIncoming);
    dateIncoming.setDate(dateIncoming.getDate() + 1);
  }
  //Logger.log("after day mod day incom: " + dateIncoming);
  return dateIncoming;
}

//calculate time to add to compute SLA
function utilitySecondConverter_(dateIncoming, dateAnswer){
  Logger.log("Incom trouble:" + dateIncoming.getDay());
  dateIncoming = resetDateWithShift_(dateIncoming);
  Logger.log("Incom trouble2:" + dateIncoming.getDay());
  var siesta = (siestaTimeStart && siestaTimeEnd);
  var sla = 0;
  do{
    sla -= siesta ? calculateSlaWithSiesta_(dateIncoming, dateAnswer): 0;
    if(dateIncoming.getDate() == dateAnswer.getDate()) {
      dateAnswer = resetDateWithShift_(dateAnswer);
      sla += Math.abs((dateAnswer.getTime() - dateIncoming.getTime())/(3600*1000));
      dateIncoming.setDate(dateIncoming.getDate() + 1);
    } else{
      sla+= Math.abs(endWorkHours - dateIncoming.getHours());
      dateIncoming.setDate(dateIncoming.getDate() + 1);
      
      dateIncoming = resetDateWithShift_(dateIncoming);
      Logger.log("Incom trouble after mod:" + dateIncoming.getDay());
     
    }
  } while(dateIncoming.getTime() < dateAnswer.getTime());
  return sla;
}

//compute the diff with sleep time
function calculateSlaWithSiesta_(dateIncoming, dateAnswer){
  Logger.log("Incoming: " + dateIncoming);
  Logger.log("Answer: " + dateAnswer);
  var totalSiesta = 0;
  if(dateIncoming.getDate() == dateAnswer.getDate()) {
      dateAnswer = resetDateWithShift_(dateAnswer);
    if(dateIncoming.getHours() < siestaTimeStart){
      if (dateAnswer.getHours() >= siestaTimeEnd){
        totalSiesta += 1;
      } else if(dateAnswer.getHours() == siestaTimeStart){
        var buffAnsDay = new Date(dateAnswer.getFullYear(), dateAnswer.getMonth(), dateAnswer.getDate());
        buffAnsDay.setHours(siestaTimeStart);
        totalSiesta += (dateAnswer.getTime() - buffAnsDay.getTime())/(60*60*1000);
      }
    } else if(dateIncoming.getHours() == siestaTimeStart){
      if (dateAnswer.getHours() >= siestaTimeEnd){
        var buffIncDay = new Date(dateIncoming.getFullYear(), dateIncoming.getMonth(), dateIncoming.getDate());
        buffIncDay.setHours(siestaTimeEnd);
        totalSiesta += (buffIncDay.getTime() - dateIncoming.getTime())/(60*60*1000);
      } else if(dateAnswer.getHours() == siestaTimeStart){
        totalSiesta += (dateAnswer.getTime() - dateIncoming.getTime())/(60*60*1000);
        }
      }
  } else{
      totalSiesta += dateIncoming.getHours() <= siestaTimeStart  ? 1:0
  }
  Logger.log("Siesta: " + totalSiesta);
  return totalSiesta  
}

//check if SLA Treshold is respected
function checkThresholdForSLA_(thresholdAnswer){
  if (thresholdAnswer < slaThreshold){
    return "SLA OK";
  } else{
    return "SLA WRONG";
  }
}

//return array with answers to the thread
function findAnswers_(messages){
  if(messages.length > 1){
   var i = 0;
   var res = [];
   for(i; i < messages.length; i++){
     var sender = stripGetFrom_(messages[i].getFrom());
     (sender.indexOf(alias) != -1) ? res.push(messages[i]):"";
   }
    return res;
  }
  return false;
}


//check day and hours for the SLA
function checkSLA_(messages, row, col){
  if(checkIncoming_(messages)){
    var answers = findAnswers_(messages);
    if(!answers){
      return false;
    }
    var from = "";
    var i = 0;
    var dateIncoming;
    var dateAnswer;
    var rowToUse = 0;
    var thresholdAnswer = 0;
    var isOutgoing = false;
    for(i; i < messages.length; i++){
      from = stripGetFrom_(messages[i].getFrom());
      //it is an answer in this block for sure! Is it?
      if (from.indexOf(alias) != -1) {
        dateAnswer = messages[i].getDate();   
        thresholdAnswer += isOutgoing ? utilitySecondConverter_(dateIncoming,dateAnswer): 0;
        thresholdAnswer = thresholdAnswer < 0 ? 0: thresholdAnswer;
        Logger.log("thresher: " + thresholdAnswer);
        writeDataInCell_(spS, checkThresholdForSLA_(thresholdAnswer), row[rowToUse], col++);
        writeDataInCell_(spS, thresholdAnswer, row[rowToUse++], col--);
        isOutgoing = false;
      } 
      if(from.indexOf(alias) == -1){
        thresholdAnswer = 0;
        dateIncoming = messages[i].getDate();
        isOutgoing = true;
      }
    }
  }  else{
    return false;
  }
}
