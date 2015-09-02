//@author: Adonis Settouf
//@mail: asettouf@lexmark.com
//ToDo: improve the cache system


//Spreadsheet to write the numbers
var spS = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//in which row we should write
var rowToWrite = 2;
//mail of the mailbox
var alias = Session.getActiveUser().getEmail();
//Labels
var labels = GmailApp.getUserLabels();
//NUmber of columns without labels
var colLen = 0;
//working hours, see the 9 and 18
var startWorkHours = 9;
var endWorkHours = 18;
//threshold for a sla to be ok (to be modified for use)
var slaThreshold = 4;
//date previous week and this week (7 days)
var arrDate = utilityDate_();
//Siesta time
var currWeek;
//Number of threads
var globalNumberOfThreads = 0;
//Number of incoming mails
var globalNumberOfMails = 0;
//Number of outgoing mails
var globalNumberOfOutgoingMails = 0;
//Number of mails sent automatically
var numberOfAutoResponseMails = 0;
//global number of columns
var globalColNum = 0;
//Agent labels
var agents = ["Geri", "Adonis","Alban","Barna"]

function clearCache(){
  var cache = CacheService.getScriptCache();
  cache.put("lastThread", 0, 3600);
  cache.put("rowToWrite", 2, 3600);
  //cache.put("globalNumberOfMails", 0, 3600);
  //cache.put("globalNumberOfThreads", 0, 3600);
}
//main function executing the the big loop
function main(){
  spS = createSpreadsheet_(spS);
  rowToWrite = init_(spS);
  var cache = CacheService.getScriptCache();
  var cachedCount = cache.get("lastThread");
  var rowCount = cache.get("rowToWrite");
  var cacheNumThread = cache.get("globalNumberOfThreads");
  var cacheNumMail = cache.get("globalNumberOfMails");
  var i = cachedCount != null ? parseInt(cachedCount) :0;
  rowToWrite = rowCount != null ? parseInt(rowCount) : rowToWrite;
  //globalNumberOfThreads = cacheNumThread != null ? parseInt(cacheNumThread) : globalNumberOfThreads;
  //globalNumberOfMails = cacheNumMail != null ? parseInt(cacheNumMail) : globalNumberOfMails;
  var query = "NOT in:draft NOT in:chats NOT in:sent after:" + arrDate[0] + " before:" + arrDate[1];
  //var query = "NOT in:draft NOT in:chats NOT in:sent after:2015/07/18 before:2015/07/26";
  Logger.log(query);
  var threads = GmailApp.search(query);
  for(i;i < threads.length; i++){
    //routine_(threads[i],threads[i].isInTrash() ? true:false);  lolz
    keepInCache_("rowToWrite", rowToWrite);
    routine_(threads[i]);
    
    //keepInCache_("globalNumberOfThreads", globalNumberOfThreads.toString());
    //keepInCache_("globalNumberOfMails", globalNumberOfMails.toString());
    keepInCache_("lastThread", i.toString());
  }
  Logger.log("Incoming: " + globalNumberOfMails);
  Logger.log("Incoming: " + globalNumberOfThreads);
  Logger.log("Number of auto mails: " + numberOfAutoResponseMails);
  Logger.log("Number of response mails: " + globalNumberOfOutgoingMails);
  writeDataInCell_(spS, globalNumberOfMails,2,globalColNum);
  writeDataInCell_(spS, globalNumberOfThreads, 2 , globalColNum + 1);
  writeDataInCell_(spS, globalNumberOfOutgoingMails - numberOfAutoResponseMails, 2 , globalColNum + 2);
  
}

//keep counter in cache for next launch
function keepInCache_(name, counter){
  var cache = CacheService.getScriptCache();
  cache.put(name, counter, 3600);
}

//utility to create new sheet each time
function createSpreadsheet_(sps){
  var curr = new Date();
  currWeek = String(parseInt(mailWeek_(curr)) - 1);
  var mailWeek = (currWeek).toString();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName().indexOf("Datas") != -1 ? "": SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setName("Datas for week: "+ mailWeek);
  if (sps.getRange(1,1).getValue() != ""){
    var sp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datas for week: "+ mailWeek) ? 
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datas for week: "+ mailWeek): SpreadsheetApp.getActiveSpreadsheet().insertSheet( "Datas for week: "+ mailWeek);
    return sp;
  }
  return SpreadsheetApp.getActiveSheet();
}

//get things started
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
  //writeDataInCell_(sps, "SLA",row,col++);
  //writeDataInCell_(sps, "Time to answer",row,col++);
  colLen = col;
   for (i = 0; i < labels.length; i++){
    writeDataInCell_(spS, labels[i].getName(),row, col++);
  }
  globalColNum = col;
  writeDataInCell_(sps, "Mail Count",1,col++);
  writeDataInCell_(sps, "Thread Count", 1 ,col++);
  writeDataInCell_(sps, "Outgoing Mail Count", 1 ,col);
  while (sps.getRange(rowToWrite, 4).getValue() != ""){ //no overwriting of existing values
    rowToWrite++;
  }
  boldingFirstRow_(sps, col + 3, row);
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
  var countSth = 0;
  var labs = thread.getLabels();
  var areLabelsPresent = findLabel_(labs);
  var mailWeek;
  var receptDay;
  for(i;i<msgs.length;i++){
    if(msgs[i].getSubject().indexOf("Attachment") > -1){
      continue;
    }
    if(msgs[i].getDate() > arrDate[0]){
      continue;
    }
    //watch out, parameters dateMail is modified through mailWeek func, not sure why though...
    dateMail = msgs[i].getDate();
    arrDayAndHour = findDayAndHour_(dateMail);
    receptDay = Utilities.formatDate(dateMail, "CET", "dd/MM/yyyy");
    //Logger.log("BC: " + dateMail);
    mailWeek = mailWeek_(dateMail);
    //Logger.log("AC: " + dateMail);
    if( mailWeek != currWeek){
      continue;
    } else{
      countSth++;
    }
    //stripGetFrom_(msgs[i].getFrom()).indexOf(alias) > -1 ? rowForSla.push(rowToWrite): "";
    col = 1;
    isIncoming = checkIncoming_Mail_(msgs[i]);
    if (isIncoming) {
      if (isIncrementMailCounter_(labs)){
        globalNumberOfMails++;
      }
      incoming ="INCOMING"; 
    } else{
       incoming = "OUTGOING";
       globalNumberOfOutgoingMails++;
    }
      
    arrToCc = checkToCc_(msgs[i]);
    type2 = msgs[i].isInTrash() ? "Junk" : checkMailType_(msgs[i], isIncoming);
    
    isUnreadTime = isUnreadTime_(msgs[i]);
    isReadString = msgs[i].isUnread() ? "Yes" : "No";
    buffRow.push(rowToWrite);
    
    writeDataInCell_(spS,receptDay, rowToWrite, col++);
    writeDataInCell_(spS, mailWeek, rowToWrite, col++);
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
  //rowForSla.length > 0 ? col = checkSLA_(msgs, rowForSla, col): "";
  var j = 0;
  var colLab;
  for(i=0;i<countSth;i++){
    colLab = colLen;
     for (j = 0; j < areLabelsPresent.length; j++){
      writeDataInCell_(spS, areLabelsPresent[j], buffRow[i], colLab++);
    }
  }
  if(countSth && !isLFMNA_(labs)){
    if (checkIncoming_(msgs)){
      Logger.log("Incoming thread: " + threadName);
      numberOfAutoResponseMails++;
    }
    if (isIncrementMailCounter_(labs)){
      globalNumberOfThreads++;
    }
  }
}

//check if thread is LFM NA
function isLFMNA_(labels){
  if(labels.length == 1 && labels[0].getName().indexOf("L F M") > -1){
    return true;
  } else{
    return false;
  }
}

//function to know if we increment global mail counter if mail is incoming and has an agent label
function isIncrementMailCounter_(labs){
  var i = 0;
  for (i; i<labs.length; i++){
    for(j=0; j < agents.length; j++){
      if (agents[j].indexOf(labs[i].getName()) > -1 ){
        return true;
      }
    }
  }
  return false;
}
//returns tab with day and hour from a date
function findDayAndHour_(date){
  res = [];
  res.push(Utilities.formatDate(date, "CET", "EEEE"));
  res.push(Utilities.formatDate(date, "CET", "H"));
  return res;
}

//find labels in a thread
function findLabel_(labs){
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
  curr.setTime(curr.getTime() + 1000*3600*24);
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
  return weekNo; 
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
  dateIncoming = resetDateWithShift_(dateIncoming);
  //var siesta = (siestaTimeStart && siestaTimeEnd);
  var sla = 0;
  do{
    //sla -= siesta ? calculateSlaWithSiesta_(dateIncoming, dateAnswer): 0;
    if(dateIncoming.getDate() == dateAnswer.getDate()) {
      Logger.log("Date answer: " + dateAnswer);
      dateAnswer = resetDateWithShift_(dateAnswer);
      Logger.log("Date answer after mod: " + dateAnswer);
      sla += Math.abs((dateAnswer.getTime() - dateIncoming.getTime())/(3600*1000));
      dateIncoming.setDate(dateIncoming.getDate() + 1);
    } else{
      sla+= Math.abs(endWorkHours - dateIncoming.getHours());
      dateIncoming.setDate(dateIncoming.getDate() + 1);
      dateIncoming = resetDateWithShift_(dateIncoming);
    }
  } while(dateIncoming.getTime() < dateAnswer.getTime());
  return sla;
}

//compute the diff with sleep time
function calculateSlaWithSiesta_(dateIncoming, dateAnswer){
  var totalSiesta = 0;
  if(dateIncoming.getDate() == dateAnswer.getDate()) {
      dateAnswer = resetDateWithShift_(dateAnswer);
    if(dateIncoming.getHours() < 12){
      if (dateAnswer.getHours() >= 13){
        totalSiesta ++;
      } else if(dateAnswer == 12){
        totalSiesta += siestaEndTime - dateAnswer.getTime()/(60*60*1000);
      }
    } else if(dateIncoming.getHours() == 12){
      if (dateAnswer.getHours() >= 13){
        totalSiesta += siestaEndTime - dateIncoming.getHours();
      } else if(dateAnswer == 12){
        totalSiesta += (dateAnswer.getTime() - dateIncoming.getTime())/(60*60*1000);
        }
      }
  } else{
      totalSiesta += dateIncoming.getHours() <= siestaTimeStart  ? 1:0
  }
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
      Logger.log(messages[i].getSubject());
      from = stripGetFrom_(messages[i].getFrom());
      //it is an answer in this block for sure! Is it?
      if (from.indexOf(alias) != -1) {
        dateAnswer = messages[i].getDate();   
        Logger.log("date incoming = " + dateIncoming);
        Logger.log("date answer = " + dateAnswer);
        thresholdAnswer += isOutgoing ? utilitySecondConverter_(dateIncoming,dateAnswer): 0;
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
    return col;
  }  else{
    return false;
  }
}
