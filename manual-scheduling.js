var PHONEINTERVIEW_ENGINEER = {
  column: 1,
  jobFunction: "Engineering"
};
var PHONEINTERVIEW_PRODUCT = {
  column: 2,
  jobFunction: "Products"
};
var PHONEINTERVIEW_BUSDEV = 3;
var PHONEINTERVIEW_CUSTSUCCES = 4;
var PHONEINTERVIEW_FRONTEND = 5;
var PHONEINTERVIEW_UX = 6;
var PHONEINTERVIEW_QA = 7;
var PHONEINTERVIEW_FINANCE = 8;
var PHONEINTERVIEW_HR = 9;
var PHONEINTERVIEW_MARKETING = 10;


//Initialize the menus on the spreadsheet UI
function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu("Actions")
  .addItem('Schedule Engineer Interview', 'manualSchedulerEngineers')
  .addItem('Schedule Product Interview', 'manualSchedulerProduct')
  .addToUi()
}

function manualSchedulerEngineers() {
   schedulePhoneInterviewsManual(PHONEINTERVIEW_ENGINEER);
}

function manualSchedulerProduct() {
   schedulePhoneInterviewsManual(PHONEINTERVIEW_PRODUCT);
}

//Schedule phone interviews manually
function schedulePhoneInterviewsManual(interviewType) {
  
  //initialize the required sheets and the ui
  var phoneInterviewSheet = SpreadsheetApp.getActive().getSheetByName("Phone Interview");
  var ui = SpreadsheetApp.getUi();
  
  //Request input from the user to determine what row to start from and how many records exist from that row
  var startRow = ui.prompt("Enter the start row").getResponseText();
  var numRecords = getLastRow(phoneInterviewSheet, 1) - startRow + 1; //Add 1 to normalize.
  
  //Request input from user on when to start the interview
  var interviewDay = ui.prompt('Enter the interview day in this format (yyyy-mm-dd)\n').getResponseText();
  var startTimeInput = ui.prompt("Enter the interview hour. (09 for 9am, 13 for 1pm etc.)\n").getResponseText();
  var interviewStartTime = new Date(interviewDay + "T" + startTimeInput + ":00:00");
  var interviewEndTime = new Date(interviewDay + "T" + startTimeInput + ":30:00")
  
  var validRowInput = false;
  
  //Determine if the user input is a number and re-prompt if otherwise
  while(!validRowInput) {
    if (isNaN(startRow)) {
      startRow = ui.prompt("Your input was not a number!\n\nEnter the start row").getResponseText();
    }
    else {
      //Get a list of all candidates and pass them to the scheduleEngineerInterviews function
      var candidateList = phoneInterviewSheet.getRange(startRow, 1, numRecords, 15).getValues();
      schedulePhoneInterviews(candidateList, interviewStartTime, interviewEndTime, phoneInterviewSheet, Number(startRow), interviewType);
      validRowInput = true;
    }
  }
  
}


//Return the email of a specified employee
function getEmployeeEmail(employeeName) {
  var employeeSheet = SpreadsheetApp.getActive().getSheetByName("Misc");
  var lastRow = getLastRow(employeeSheet, 19);
  var employeeMailList = employeeSheet.getRange(2, 19, lastRow, 2).getValues();
  for (var i in employeeMailList) {
    if (employeeName == employeeMailList[i][0]) {
      return employeeMailList[i][1];
    }
  }
  return null;
}


//Schedule phone interviews for engineering candidates
function schedulePhoneInterviews(candidates, interviewStartTime, interviewEndTime, phoneInterviewSheet, startRow, interviewType) {
  //initialize the phone interviewers spreadsheet, the number of scheduled interviewers and get the total number of interviewers
  var phoneInterviewersSheet = SpreadsheetApp.getActive().getSheetByName("Phone Interviewers");
  var numInterviewerRows = getLastRow(phoneInterviewersSheet, interviewType.column);
  var numberOfInterviewers = numInterviewerRows - 1; //subtract 1 to remove header
  var scheduledInterviews = 0;
  
  //Get the list of engineering phone interviewers
  var interviewers = getInterviewers(interviewType.column);
  
  for (var i in candidates) {
    //Use pattern to check if candidate is an engineer
    if (candidates[i][13] == interviewType.jobFunction) {
      var candidateEmail = candidates[i][1];
      var candidatePhone = candidates[i][4];
      var candidateCV = candidates[i][5];
      var interviewer = interviewers[scheduledInterviews];
      var interviewerEmail = getEmployeeEmail(interviewer);
      sendPhoneInterviewInvite(candidateEmail, candidatePhone, candidateCV, interviewerEmail, interviewStartTime, interviewEndTime);
      var row = Number(startRow) + Number(i);
      phoneInterviewSheet.getRange(row, 10, 1, 2).setValues([[interviewer, interviewStartTime]]);
      scheduledInterviews++
    }
    
    //Check to ensure number of scheduled interviews has not exceeded number of interviewers
    if (scheduledInterviews >= numberOfInterviewers) {
      break;
    }
    
  }
  
  return;
}


//Send phone interview invite
function sendPhoneInterviewInvite(candidateEmail, candidatePhone, candidateCV, interviewerEmail, startTime, endTime) {
  
  var guestEmails = candidateEmail + "," + interviewerEmail;
  
  CalendarApp.getCalendarById("recruitment@teamapt.com").createEvent('TeamApt Limited - Phone Interview', startTime, endTime, {
    description: 'CV: ' + candidateCV + '\n\nPhone: ' + candidatePhone,
    guests: guestEmails
  });
  
}


//return the last row in a sheet 
//function getLastRow(sheet, colToCheck){
//  var lastRow = sheet.getMaxRows();
//  var values = sheet.getRange(1, colToCheck, lastRow).getValues();
//  
//  while(values[lastRow - 1] == "" && lastRow > 0){
//    lastRow--;
//  }
//  
//  return lastRow;
//}


function test() {
  var emails = "amiratobi@yahoo.com" + "," + "amiratobi@gmail.com";
  var startTime = new Date(2018, 11, 30, 10, 30);
  var endTime = new Date(2018, 11, 30, 11, 00);
  CalendarApp.getCalendarById("recruitment@teamapt.com").createEvent('Test', startTime, endTime, {
    guests: emails
  });
}


//return the list of interviewers for a job function
function getInterviewers(jobFunction) {
  var phoneInterviewersSheet = SpreadsheetApp.getActive().getSheetByName("Phone Interviewers");
  var numRows = getLastRow(phoneInterviewersSheet, jobFunction);
  var interviewerList = phoneInterviewersSheet.getRange(2, jobFunction, numRows).getValues();
  var interviewers = []
  
  for (var i in interviewerList) {
    if (interviewerList[i] != "") {
      interviewers.push(interviewerList[i][0]);   
    }
  } 
  
  return interviewers;
}