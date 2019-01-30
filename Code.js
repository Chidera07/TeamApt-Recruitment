//Sheets
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var prospectiveFormResponses = spreadsheet.getSheetByName("Prospectives");
var portfolioReviewSheet = spreadsheet.getSheetByName("Portfolio Review");
var phoneInterviewSheet = spreadsheet.getSheetByName("Phone Interview");
var firstTechnicalnterviewSheet = spreadsheet.getSheetByName("1st Technical Review");
var technicalInterviewSheet = spreadsheet.getSheetByName("Technical Interview");
var physicalInterviewSheet = spreadsheet.getSheetByName("Physical Interview");
var hiredSheet = spreadsheet.getSheetByName("Accepted");
var failedSheet = spreadsheet.getSheetByName("Failed");
var waitingSheet = spreadsheet.getSheetByName("Waiting List");
var tatSheet = spreadsheet.getSheetByName("TAT");
var summarySheet = spreadsheet.getSheetByName("Daily Summary");
var interviewerSheet = spreadsheet.getSheetByName("Phone Interviewers/ Portfolio Reviewers");
var technicalInterviewerSheet = spreadsheet.getSheetByName("1st and 2nd Technical Interviewers");
var physicalInterviewerSheet = spreadsheet.getSheetByName("Physical Interviewers");
var internAppSheet = spreadsheet.getSheetByName("Internship Application");
var scholarAppSheet = spreadsheet.getSheetByName("Scholarship Application");
var oldCompletedSheet = spreadsheet.getSheetByName("Completed");
var decisionMatrixSheet = spreadsheet.getSheetByName("Decision Matrix");
var vpdecisionSheet = spreadsheet.getSheetByName("VP Decision Sheet");
var summarySheet = spreadsheet.getSheetByName("Interviewers Summary");
var interviewMatrixSheet = spreadsheet.getSheetByName("Interview Matrix");

//Columns
var TIMESTAMPCOLUMN = 1;
var TATNAMECOLUMN = 1;
var INTERVIEWERNAMECOLUMN = 1;
var EMAILADDRESSCOLUMN = 2;
var TATROLECOLUMN = 2;
var INTERVIEWERFUNCTIONSCOLUMN = 2;
var TATPROSPECTIVESCOLUMN = 3;
var NAMECOLUMN = 3;
var INTERVIEWERCALLTIMECOLUMN = 3;
var ROLECOLUMN = 4;
var TATPHONECOLUMN = 4;
var INTERVIEWERCALLSTAKENCOLUMN = 4;
var TATTECHNICALCOLUMN = 5;
var TATPHYSICALCOLUMN = 6;
var CVCOLUMN = 6;
var PORTFOLIOCOLUMN = 7;
var PROSPECTIVESSKYPEIDCOLUMN = 10;
var PHONEINTERVIEWERCOLUMN = 13;
var PHONECALLTIMECOLUMN = 14;
var PHONESKYPEIDCOLUMN = 10;
var PROSPECTIVESDECISIONCOLUMN = 11;
var PORTFOLIOREVIEWERCOLUMN = 11;
var PROSPECTIVESFUNCTIONCOLUMN = 12;
var PORTFOLIOREVIEWCOMMENTCOLUMN = 12;
var PORTFOLIOREVIEWDECISIONCOLUMN = 13;
var PHONEINTERVIEWERCOMMENTCOLUMN = 15;
var PHONEDECISIONCOLUMN = 16;
var TECHNICALINTERVIEWERCOLUMN = 16;
var FIRSTTECHNICALINTERVIEWERCOLUMN = 16;
var FIRSTTECHNICALCALLTIMECOLUMN = 17;
var PHONEFUNCTIONCOLUMN = 17;
var FIRSTTECHNICALINTERVIEWERCOMMENTCOLUMN = 18;
var TECHNICALCALLTIMECOLUMN = 20;
var FIRSTTECHNICALDECISIONCOLUMN = 20;
var PHYSICALCALLTIMECOLUMN = 21;
var TECHNICALINTERVIEWERCOMMENTCOLUMN = 21;
var TECHNICALDECISIONCOLUMN = 23;
var PHYSICALINTERVIEWERCOLUMN = 25;
var PHYSICALINTERVIEWERCOMMENTCOLUMN = 26;
var PHYSICALINTERVIEWERTWOCOLUMN = 27;
var WAITINGLISTSTATUSCOLUMN = 31;
var FAILEDSTATUSCOLUMN = 33;
var PHYSICALDECISIONCOLUMN = 33;
var FAILEDSTAGECOLUMN = 34;

var functions = ["Engineering", "Products","Network & Infrastructure", "Customer Success", "Marketing", 
                 "Finance", "Legal & Risk", "Business Development", "HR and Admin", "Management", "MIS", "Creatives"];

function onFormSubmit(e){
  var eRange = e.source.getActiveRange();
  var sheet = eRange.getSheet().getName();
  var row = eRange.getRow();
  
  if(sheet === "Prospectives"){
    prospectiveFormResponses.getRange(row, PROSPECTIVESDECISIONCOLUMN).setValue("Yet to contact");
    var name = prospectiveFormResponses.getRange(row, NAMECOLUMN).getValue();
    var emailAddress = prospectiveFormResponses.getRange(row, EMAILADDRESSCOLUMN).getValue();
    var duplicate = isDuplicate(emailAddress,prospectiveFormResponses, EMAILADDRESSCOLUMN);
  
    if(!isNaN(duplicate)){
      prospectiveFormResponses.deleteRow(row);
    }
    else{
      sendReceiptEmail(name,emailAddress);
    }
  }
  if(sheet === "Internship Application"){
    sendNotificationEmail("internship");
  }
  if(sheet === "Scholarship Application"){
    sendNotificationEmail("scholarship");
  }
}
//***********************************************************************************************************************************************************************************
function check(role){
  var checker = {};
  var values;
  var lastRow = getLastRow(interviewMatrixSheet, 1);
  var roles = interviewMatrixSheet.getRange(2, 1, lastRow, 1).getValues();
  for(var i = 2; i <= lastRow; i++){
    if(roles[i-2][0] === role){
      values = interviewMatrixSheet.getRange(i, 2, 1, 5).getValues();
    }
  }
  checker.portfolio = values[0][0];
  checker.phone = values[0][1];
  checker.technical1 = values[0][2];
  checker.technical2 = values[0][3];
  checker.physical = values[0][4];
  
  return checker;
}
//***********************************************************************************************************************************************************************************
function onEdit(e){
  var ui = SpreadsheetApp.getUi();
  var eRange = e.source.getActiveRange();
  var sheet = eRange.getSheet();
  var sheetName = sheet.getName();
  var row = eRange.getRow();
  var column = eRange.getColumn();
  var status = e.value;
  var role = sheet.getRange(row, ROLECOLUMN).getValue();
  var checker = check(role);
  var candidateFunction = sheet.getRange(row, PHONEFUNCTIONCOLUMN).getValue();
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Prospectives" && column === PROSPECTIVESDECISIONCOLUMN && (status === "Awaiting phone Interview" ||
    status === "Failed" || status === "Add to waiting list" || status === "Awaiting portfolio review") 
    && prospectiveFormResponses.getRange(row, TIMESTAMPCOLUMN).getValue() === ""){
      ui.alert("A timestamp needs to be assigned before moving this candidate");
      prospectiveFormResponses.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Prospectives" && column === PROSPECTIVESDECISIONCOLUMN && (status === "Awaiting phone Interview" ||
    status === "Failed" || status === "Add to waiting list" || status === "Awaiting portfolio review") 
    && prospectiveFormResponses.getRange(row, EMAILADDRESSCOLUMN).getValue() === ""){
      ui.alert("An email address needs to be assigned before moving this candidate");
      prospectiveFormResponses.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Prospectives" && column === PROSPECTIVESDECISIONCOLUMN && status === "Awaiting phone Interview"
    && checker.phone === "No"){
      ui.alert("This job role cannot be moved to phone interview stage");
      prospectiveFormResponses.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Prospectives" && column === PROSPECTIVESDECISIONCOLUMN && status === "Awaiting portfolio review"
    && prospectiveFormResponses.getRange(row, PORTFOLIOCOLUMN).getValue() === ""){
      ui.alert("Candidate needs to have a portfolio");
      prospectiveFormResponses.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Prospectives" && column === PROSPECTIVESDECISIONCOLUMN && status === "Awaiting portfolio review" && checker.portfolio === "No"){
      ui.alert("This job role cannot be moved to portfolio review stage");
      prospectiveFormResponses.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Prospectives" && column === PROSPECTIVESDECISIONCOLUMN && status === "Awaiting phone Interview" && checker.phone === "Yes"){
    var result = ui.alert("Are you sure you want to move this candidate to the Phone Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = prospectiveFormResponses.getRange(row, TIMESTAMPCOLUMN, 1, 12).getValues();
      var phoneInterviewSheetLastRow = getLastRow(phoneInterviewSheet, TIMESTAMPCOLUMN);
      prospectiveFormResponses.getRange(row, EMAILADDRESSCOLUMN, 1, 9).copyTo(phoneInterviewSheet.getRange(phoneInterviewSheetLastRow+1, 2));
      phoneInterviewSheet.getRange(phoneInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      var schedule = scheduler("Phone", candidateDetails[0][11], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      phoneInterviewSheet.getRange(phoneInterviewSheetLastRow+1, PHONEINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      phoneInterviewSheet.getRange(phoneInterviewSheetLastRow+1, PHONECALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Prospectives", candidateDetails[0][0]);
      prospectiveFormResponses.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      prospectiveFormResponses.getRange(row, column).setValue(e.oldValue);
    } 
  }
//**********************************************************************************************************************************************************************************
  if(sheetName === "Prospectives" && column === PROSPECTIVESDECISIONCOLUMN && status === "Failed"){
    var result = ui.alert("Are you sure you want to end the application process for this candidate?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = prospectiveFormResponses.getRange(row, TIMESTAMPCOLUMN, 1, 11).getValues();
      var failedSheetLastRow = getLastRow(failedSheet, TIMESTAMPCOLUMN);
      prospectiveFormResponses.getRange(row, EMAILADDRESSCOLUMN, 1, 8).copyTo(failedSheet.getRange(failedSheetLastRow+ 1, 2));
      failedSheet.getRange(failedSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTATUSCOLUMN).setValue("Failed");
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTAGECOLUMN).setValue("Prospectives");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Prospectives", candidateDetails[0][0]);
      prospectiveFormResponses.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      prospectiveFormResponses.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//**********************************************************************************************************************************************************************************
  if(sheetName === "Prospectives" && column === 11 && status === "Add to waiting list") {
    var result = ui.alert("Are you sure you want to move this candidate to the waiting list?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = prospectiveFormResponses.getRange(row, TIMESTAMPCOLUMN, 1, 11).getValues();
      var waitingListSheetLastRow = getLastRow(waitingSheet, TIMESTAMPCOLUMN);
      prospectiveFormResponses.getRange(row, EMAILADDRESSCOLUMN, 1, 9).copyTo(waitingSheet.getRange(waitingListSheetLastRow+ 1, 2));
      waitingSheet.getRange(waitingListSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      waitingSheet.getRange(waitingListSheetLastRow+1, WAITINGLISTSTATUSCOLUMN).setValue("Waiting List");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Prospectives", candidateDetails[0][0]);
      prospectiveFormResponses.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      prospectiveFormResponses.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Prospectives" && column === 11 && status === "Awaiting portfolio review" && checker.portfolio === "Yes") {
    var result = ui.alert("Move candidate for portfolio review?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = prospectiveFormResponses.getRange(row, TIMESTAMPCOLUMN, 1, 12).getValues();
      var portfolioReviewSheetLastRow = getLastRow(portfolioReviewSheet, TIMESTAMPCOLUMN);
      prospectiveFormResponses.getRange(row, EMAILADDRESSCOLUMN, 1, 9).copyTo(portfolioReviewSheet.getRange(portfolioReviewSheetLastRow+ 1, 2));
      portfolioReviewSheet.getRange(portfolioReviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      var schedule = scheduler("Portfolio Review", candidateDetails[0][11], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5], candidateDetails[0][6]);
      portfolioReviewSheet.getRange(portfolioReviewSheetLastRow+1, PORTFOLIOREVIEWERCOLUMN).setValue(schedule.interviewer);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Prospectives", candidateDetails[0][0]);
      prospectiveFormResponses.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      prospectiveFormResponses.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && (status === "YES" || status === "NO" || status === "MAYBE" || status === "WAITING LIST")
    && portfolioReviewSheet.getRange(row, PORTFOLIOREVIEWCOMMENTCOLUMN).getValue() === ""){
      ui.alert("A comment needs to be assigned before moving this candidate");
      portfolioReviewSheet.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && (status === "YES" || status === "NO" || status === "MAYBE" || status === "WAITING LIST")
    && portfolioReviewSheet.getRange(row, TIMESTAMPCOLUMN).getValue() === ""){
      ui.alert("A timestamp needs to be assigned before moving this candidate");
      portfolioReviewSheet.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && (status === "YES" || status === "NO" || status === "MAYBE" || status === "WAITING LIST") 
    && portfolioReviewSheet.getRange(row, EMAILADDRESSCOLUMN).getValue() === ""){
      ui.alert("An email address needs to be assigned before moving this candidate");
      portfolioReviewSheet.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && (status === "YES" || status === "MAYBE") && checker.phone === "Yes"){
    var result = ui.alert("Are you sure you want to move this candidate to the Phone Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = portfolioReviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 14).getValues();
      var phoneInterviewSheetLastRow = getLastRow(phoneInterviewSheet, TIMESTAMPCOLUMN);
      portfolioReviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 11).copyTo(phoneInterviewSheet.getRange(phoneInterviewSheetLastRow+1, 2));
      phoneInterviewSheet.getRange(phoneInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      var schedule = scheduler("Phone", candidateDetails[0][13], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      phoneInterviewSheet.getRange(phoneInterviewSheetLastRow+1, PHONEINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      phoneInterviewSheet.getRange(phoneInterviewSheetLastRow+1, PHONECALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Portfolio Review", candidateDetails[0][0]);
      calculateDailySummary("Porfolio Review");
      portfolioReviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      portfolioReviewSheet.getRange(row, column).setValue(e.oldValue);
    }
    return;
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && (status === "YES" || status === "MAYBE") && checker.technical1 === "Yes"){
    var result = ui.alert("Are you sure you want to move this candidate to the First Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = portfolioReviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 14).getValues();
      var firstTechnicalInterviewSheetLastRow = getLastRow(firstTechnicalnterviewSheet, TIMESTAMPCOLUMN);
      portfolioReviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 11).copyTo(firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, 2));
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      var schedule = scheduler("First Technical", candidateDetails[0][13], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, FIRSTTECHNICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, FIRSTTECHNICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Portfolio Review", candidateDetails[0][0]);
      calculateDailySummary("Porfolio Review");
      portfolioReviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      portfolioReviewSheet.getRange(row, column).setValue(e.oldValue);
    }
    return;
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && (status === "YES" || status === "MAYBE") && checker.technical2 === "Yes"){
    var result = ui.alert("Are you sure you want to move this candidate to the Second Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = portfolioReviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 14).getValues();
      var technicalInterviewSheetLastRow = getLastRow(technicalInterviewSheet, TIMESTAMPCOLUMN);
      portfolioReviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 11).copyTo(technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, 2));
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      var schedule = scheduler("Technical", candidateDetails[0][13], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TECHNICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TECHNICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Portfolio Review", candidateDetails[0][0]);
      calculateDailySummary("Porfolio Review");
      portfolioReviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      portfolioReviewSheet.getRange(row, column).setValue(e.oldValue);
    }
    return;
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && (status === "YES" || status === "MAYBE") && checker.physical === "Yes"){
    var result = ui.alert("Are you sure you want to move this candidate to the Physical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = portfolioReviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 14).getValues();
      var physicalInterviewSheetLastRow = getLastRow(physicalInterviewSheet, TIMESTAMPCOLUMN);
      portfolioReviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 11).copyTo(physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 2));
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      var schedule = scheduler("Physical", candidateDetails[0][13], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERTWOCOLUMN).setValue(schedule.interviewer2);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Portfolio Review", candidateDetails[0][0]);
      calculateDailySummary("Porfolio Review");
      portfolioReviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      portfolioReviewSheet.getRange(row, column).setValue(e.oldValue);
    }
    return;
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && status === "NO"){
    var result = ui.alert("Are you sure you want to end the application process for this candidate?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = portfolioReviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 14).getValues();
      var failedSheetLastRow = getLastRow(failedSheet, TIMESTAMPCOLUMN);
      portfolioReviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 11).copyTo(failedSheet.getRange(failedSheetLastRow+ 1, 2));
      failedSheet.getRange(failedSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTATUSCOLUMN).setValue("Failed");
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTAGECOLUMN).setValue("Prospectives");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Portfolio Review", candidateDetails[0][0]);
      calculateDailySummary("Porfolio Review");
      portfolioReviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      portfolioReviewSheet.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Portfolio Review" && column === PORTFOLIOREVIEWDECISIONCOLUMN && status === "WAITING LIST"){
    var result = ui.alert("Are you sure you want to move this candidate to the waiting list?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = portfolioReviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 14).getValues();
      var waitingListSheetLastRow = getLastRow(waitingSheet, TIMESTAMPCOLUMN);
      portfolioReviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 11).copyTo(waitingSheet.getRange(waitingListSheetLastRow+ 1, 2));
      waitingSheet.getRange(waitingListSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      waitingSheet.getRange(waitingListSheetLastRow+1, WAITINGLISTSTATUSCOLUMN).setValue("Waiting List");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Portfolio Review", candidateDetails[0][0]);
      calculateDailySummary("Porfolio Review");
      portfolioReviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      portfolioReviewSheet.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && (status === "YES" || status === "MAYBE" || status === "NO" || status === "WAITING LIST") 
    && (phoneInterviewSheet.getRange(row, PHONEINTERVIEWERCOMMENTCOLUMN).getValue() === "")) {
    ui.alert("Please enter a comment before moving candidate");
    phoneInterviewSheet.getRange(row, column).setValue(e.oldValues);
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && (status === "YES" || status === "NO" || status === "MAYBE" || status === "WAITING LIST")
    && phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN).getValue() === ""){
      ui.alert("A timestamp needs to be assigned before moving this candidate");
      phoneInterviewSheet.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && (status === "YES" || status === "NO" || status === "MAYBE" || status === "WAITING LIST") 
    && phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN).getValue() === ""){
      ui.alert("An email address needs to be assigned before moving this candidate");
      phoneInterviewSheet.getRange(row,  column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************
  //four
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && status === "YES" && checker.technical1 === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the First Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 17).getValues();
      var firstTechnicalInterviewSheetLastRow = getLastRow(firstTechnicalnterviewSheet, TIMESTAMPCOLUMN);
      phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 14).copyTo(firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, 2));
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, 19).setValue("YES");
//      var schedule = scheduler("First Technical", candidateDetails[0][16], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, FIRSTTECHNICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, FIRSTTECHNICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Phone Interview", candidateDetails[0][0]);
      moveToTechnicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
//      decisionMatrix("phone", row);
      calculateDailySummary("Phone Interview");
      phoneInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      phoneInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  //five
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && status === "MAYBE" && checker.technical1 === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the First Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 17).getValues();
      var firstTechnicalInterviewSheetLastRow = getLastRow(firstTechnicalnterviewSheet, TIMESTAMPCOLUMN);
      phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 14).copyTo(firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, 2));
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, 19).setValue("MAYBE");
//      var schedule = scheduler("First Technical", candidateDetails[0][16], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, FIRSTTECHNICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      firstTechnicalnterviewSheet.getRange(firstTechnicalInterviewSheetLastRow+1, FIRSTTECHNICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Phone Interview", candidateDetails[0][0]);
      moveToTechnicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
//      decisionMatrix("phone", row);
      calculateDailySummary("Phone Interview");
      phoneInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      phoneInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  //six
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && status === "YES" && checker.technical2 === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the First Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 17).getValues();
      var technicalInterviewSheetLastRow = getLastRow(technicalInterviewSheet, TIMESTAMPCOLUMN);
      phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 14).copyTo(technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, 2));
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, 22).setValue("YES");
      var schedule = scheduler("Technical", candidateDetails[0][16], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TECHNICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TECHNICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Phone Interview", candidateDetails[0][0]);
      moveToTechnicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
//      decisionMatrix("phone", row);
      calculateDailySummary("Phone Interview");
      phoneInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      phoneInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  //seven
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && status === "MAYBE" && checker.technical2 === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the First Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 17).getValues();
      var technicalInterviewSheetLastRow = getLastRow(technicalInterviewSheet, TIMESTAMPCOLUMN);
      phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 14).copyTo(technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, 2));
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, 22).setValue("MAYBE");
      var schedule = scheduler("Technical", candidateDetails[0][16], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TECHNICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TECHNICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Phone Interview", candidateDetails[0][0]);
      moveToTechnicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
//      decisionMatrix("phone", row);
      calculateDailySummary("Phone Interview");
      phoneInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      phoneInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  //eight
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && status === "YES" && checker.physical === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the First Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 17).getValues();
      var physicalInterviewSheetLastRow = getLastRow(physicalInterviewSheet, TIMESTAMPCOLUMN);
      phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 14).copyTo(physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 2));
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 22).setValue("YES");
      var schedule = scheduler("Physical", candidateDetails[0][16], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERTWOCOLUMN).setValue(schedule.interviewer2);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Phone Interview", candidateDetails[0][0]);
      moveToPhysicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
//      decisionMatrix("phone", row);
      calculateDailySummary("Phone Interview");
      phoneInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      phoneInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  //nine
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && status === "MAYBE" && checker.physical === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the First Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 17).getValues();
      var physicalInterviewSheetLastRow = getLastRow(physicalInterviewSheet, TIMESTAMPCOLUMN);
      phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 14).copyTo(physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 2));
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 22).setValue("MAYBE");
      var schedule = scheduler("Physical", candidateDetails[0][16], candidateDetails[0][1], candidateDetails[0][4], candidateDetails[0][5]);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERTWOCOLUMN).setValue(schedule.interviewer2);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Phone Interview", candidateDetails[0][0]);
      moveToPhysicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
//      decisionMatrix("phone", row);
      calculateDailySummary("Phone Interview");
      phoneInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      phoneInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && status === "NO") {
    var result = ui.alert("Are you sure you want to end the application process for this candidate?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 17).getValues();
      var failedSheetLastRow = getLastRow(failedSheet, 1);
      phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 14).copyTo(failedSheet.getRange(failedSheetLastRow+1, 2));
      failedSheet.getRange(failedSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTATUSCOLUMN).setValue("Failed");
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTAGECOLUMN).setValue("Phone Interview");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Phone Interview", candidateDetails[0][0]);
      sendRejectEmail_Phone(candidateDetails[0][2], candidateDetails[0][1]);
      calculateDailySummary("Phone Interview");
      phoneInterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      phoneInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************   
  if(sheetName === "Phone Interview" && column === PHONEDECISIONCOLUMN && status === "WAITING LIST") {
    var result = ui.alert("Are you sure you want to move this candidate to the waiting list?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = phoneInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 17).getValues();
      var waitingListSheetLastRow = getLastRow(waitingSheet, 1);
      phoneInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 14).copyTo(waitingSheet.getRange(waitingListSheetLastRow+ 1, 2));
      waitingSheet.getRange(waitingListSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      waitingSheet.getRange(waitingListSheetLastRow+1, WAITINGLISTSTATUSCOLUMN).setValue("Waiting List");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Phone Interview", candidateDetails[0][0]);
      calculateDailySummary("Phone Interview");
      phoneInterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      phoneInterviewSheet.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************  
  if(sheetName === "1st Technical Review" && column === FIRSTTECHNICALDECISIONCOLUMN && (status === "YES" || status === "MAYBE" || status === "NO" || status === "WAITING LIST") 
    && (firstTechnicalnterviewSheet.getRange(row, FIRSTTECHNICALINTERVIEWERCOMMENTCOLUMN).getValue() === "")) {
    ui.alert("Please enter a comment before moving candidate");
    firstTechnicalnterviewSheet.getRange(row, column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************    
  //ten
  if(sheetName === "1st Technical Review" && column === FIRSTTECHNICALDECISIONCOLUMN && (status === "YES" || status === "MAYBE") && checker.technical2 === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the Second Technical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = firstTechnicalnterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 21).getValues();
      var technicalInterviewSheetLastRow = getLastRow(technicalInterviewSheet, 1);
      firstTechnicalnterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 17).copyTo(technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, 2));
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, 22).setValue(candidateDetails[0][18]);
      var schedule = scheduler("Technical", candidateDetails[0][20], candidateDetails[0][0], candidateDetails[0][3], candidateDetails[0][4]);
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TECHNICALCALLTIMECOLUMN).setValue(schedule.callTime);
      technicalInterviewSheet.getRange(technicalInterviewSheetLastRow+1, TECHNICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "First Technical Interview", candidateDetails[0][0]);
      moveToTechnicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
      //decisionMatrix("first technical", row);
      calculateDailySummary("First Technical Interview");
      firstTechnicalnterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      firstTechnicalnterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************    
  //eleven
  if(sheetName === "1st Technical Review" && column === FIRSTTECHNICALDECISIONCOLUMN && (status === "YES" || status === "MAYBE") && checker.physical === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the Physical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = firstTechnicalnterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 21).getValues();
      var physicalInterviewSheetLastRow = getLastRow(physicalInterviewSheet, 1);
      firstTechnicalnterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 17).copyTo(physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 2));
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 22).setValue(candidateDetails[0][18]);
      var schedule = scheduler("Physical", candidateDetails[0][20], candidateDetails[0][0], candidateDetails[0][3], candidateDetails[0][4]);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERTWOCOLUMN).setValue(schedule.interviewer2);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "First Technical Interview", candidateDetails[0][0]);
      moveToPhysicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
      //decisionMatrix("first technical", row);
      calculateDailySummary("First Technical Interview");
      firstTechnicalnterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      firstTechnicalnterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************    
  if(sheetName === "1st Technical Review" && column === FIRSTTECHNICALDECISIONCOLUMN && status === "NO") {
    var result = ui.alert("Are you sure you want to end the application process for this candidate?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = firstTechnicalnterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 21).getValues();
      var failedSheetLastRow = getLastRow(failedSheet, 1);
      firstTechnicalnterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 17).copyTo(failedSheet.getRange(failedSheetLastRow+1, 2));
      failedSheet.getRange(failedSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTATUSCOLUMN).setValue("Failed");
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTAGECOLUMN).setValue("1st Technical Review");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "First Technical Interview", candidateDetails[0][0]);
      sendRejectEmail_Technical(candidateDetails[0][2], candidateDetails[0][1]);
      calculateDailySummary("First Technical Interview");
      firstTechnicalnterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      firstTechnicalnterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************    
  if(sheetName === "1st Technical Review" && column === FIRSTTECHNICALDECISIONCOLUMN && status === "WAITING LIST") {
    var result = ui.alert("Are you sure you want to move this candidate to the waiting list?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = firstTechnicalnterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 21).getValues();
      var waitingListSheetLastRow = getLastRow(waitingSheet, 1);
      firstTechnicalnterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 17).copyTo(waitingSheet.getRange(waitingListSheetLastRow+ 1, 2));
      waitingSheet.getRange(waitingListSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      waitingSheet.getRange(waitingListSheetLastRow+1, WAITINGLISTSTATUSCOLUMN).setValue("Waiting List");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "First Technical Interview", candidateDetails[0][0]);
      calculateDailySummary("First Technical Interview");
      firstTechnicalnterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      firstTechnicalnterviewSheet.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************     
  if(sheetName === "Technical Interview" && column === TECHNICALDECISIONCOLUMN && (status === "YES" || status === "MAYBE" || status === "NO" || status === "WAITING LIST") 
    && (technicalInterviewSheet.getRange(row, TECHNICALINTERVIEWERCOMMENTCOLUMN).getValue() === "")) {
    ui.alert("Please enter a comment before moving candidate");
    technicalInterviewSheet.getRange(row, column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************    
  //twelve
  if(sheetName === "Technical Interview" && column === TECHNICALDECISIONCOLUMN && status === "YES" && checker.physical === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the Physical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = technicalInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 24).getValues();
      var physicalInterviewSheetLastRow = getLastRow(physicalInterviewSheet, 1);
      technicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 20).copyTo(physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 2));
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 22).setValue(candidateDetails[0][21]);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 23).setValue("YES");
      var schedule = scheduler("Physical", candidateDetails[0][23], candidateDetails[0][0], candidateDetails[0][3], candidateDetails[0][4]);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERTWOCOLUMN).setValue(schedule.interviewer2);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Technical Interview", candidateDetails[0][0]);
      moveToPhysicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
      //decisionMatrix("first technical", row);
      calculateDailySummary("Technical Interview");
      technicalInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      technicalInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************    
  //thirteen
  if(sheetName === "Technical Interview" && column === TECHNICALDECISIONCOLUMN && status === "MAYBE" && checker.physical === "Yes") {
    var result = ui.alert("Are you sure you want to move this candidate to the Physical Interview stage?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = technicalInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 24).getValues();
      var physicalInterviewSheetLastRow = getLastRow(physicalInterviewSheet, 1);
      technicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 20).copyTo(physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 2));
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 22).setValue(candidateDetails[0][21]);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, 23).setValue("MAYBE");
      var schedule = scheduler("Physical", candidateDetails[0][23], candidateDetails[0][0], candidateDetails[0][3], candidateDetails[0][4]);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERCOLUMN).setValue(schedule.interviewer);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALINTERVIEWERTWOCOLUMN).setValue(schedule.interviewer2);
      physicalInterviewSheet.getRange(physicalInterviewSheetLastRow+1, PHYSICALCALLTIMECOLUMN).setValue(schedule.callTime);
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Technical Interview", candidateDetails[0][0]);
      moveToPhysicalInterviewEmail(candidateDetails[0][2], candidateDetails[0][1]);
      //decisionMatrix("first technical", row);
      calculateDailySummary("Technical Interview");
      technicalInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      technicalInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************    
 if(sheetName === "Technical Interview" && column === TECHNICALDECISIONCOLUMN && status === "NO") {
    var result = ui.alert("Are you sure you want to end the application process for this candidate?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = technicalInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 24).getValues();
      var failedSheetLastRow = getLastRow(failedSheet, 1);
      technicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 20).copyTo(failedSheet.getRange(failedSheetLastRow+1, 2));
      failedSheet.getRange(failedSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTATUSCOLUMN).setValue("Failed");
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTAGECOLUMN).setValue("1st Technical Review");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Technical Interview", candidateDetails[0][0]);
      sendRejectEmail_Technical(candidateDetails[0][2], candidateDetails[0][1]);
      calculateDailySummary("Technical Interview");
      technicalInterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      technicalInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************    
  if(sheetName === "Technical Interview" && column === TECHNICALDECISIONCOLUMN && status === "WAITING LIST") {
    var result = ui.alert("Are you sure you want to move this candidate to the waiting list?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = technicalInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 24).getValues();
      var waitingListSheetLastRow = getLastRow(waitingSheet, 1);
      technicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 20).copyTo(waitingSheet.getRange(waitingListSheetLastRow+ 1, 2));
      waitingSheet.getRange(waitingListSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      waitingSheet.getRange(waitingListSheetLastRow+1, WAITINGLISTSTATUSCOLUMN).setValue("Waiting List");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Technical Interview", candidateDetails[0][0]);
      calculateDailySummary("Technical Interview");
      technicalInterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      technicalInterviewSheet.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************      
  if(sheetName === "Physical Interview" && column === PHYSICALDECISIONCOLUMN && (status === "YES" || status === "MAYBE" || status === "NO" || status === "WAITING LIST") 
    && (physicalInterviewSheet.getRange(row, PHYSICALINTERVIEWERCOMMENTCOLUMN).getValue() === "")) {
    ui.alert("Please enter a comment before moving candidate");
    physicalInterviewSheet.getRange(row, column).setValue(e.oldValue);
  }
//***********************************************************************************************************************************************************************************    
  if(sheetName === "Physical Interview" && column === PHYSICALDECISIONCOLUMN && status === "YES") {
    var result = ui.alert("Are you sure you want to hire this candidate?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = physicalInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 34).getValues();
      var hiredSheetLastRow = getLastRow(hiredSheet, 1);
      var decision = makeFinalDecision(hiredSheetLastRow+1, candidateDetails[0][21], candidateDetails[0][22], "YES");
      if(decision === "FULL TIME HIRE"){
        physicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(hiredSheet.getRange(hiredSheetLastRow+ 1, 2));
        hiredSheet.getRange(hiredSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
        hiredSheet.getRange(hiredSheetLastRow+1, 33).setValue("YES");
        hiredSheet.getRange(hiredSheetLastRow+1, 34).setValue("FULL TIME HIRE");
        hiredSheet.getRange(hiredSheetLastRow+1, 35).setValue("Awaiting Feedback");
      }
      else if(decision === "CONTRACT"){
        physicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(hiredSheet.getRange(hiredSheetLastRow+ 1, 2));
        hiredSheet.getRange(hiredSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
        hiredSheet.getRange(hiredSheetLastRow+1, 33).setValue("YES");
        hiredSheet.getRange(hiredSheetLastRow+1, 34).setValue("CONTRACT");
        hiredSheet.getRange(hiredSheetLastRow+1, 35).setValue("Awaiting Feedback");
      }
      else{
        var vpdecisionSheetLastRow = getLastRow(vpdecisionSheet, 1);
        physicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(vpdecisionSheet.getRange(vpdecisionSheetLastRow+ 1, 2));
        vpdecisionSheet.getRange(vpdecisionSheetLastRow+1, 1).setValue(new Date());
        vpdecisionSheet.getRange(vpdecisionSheetLastRow+1, 33).setValue("YES");
        sendVPEMails();
      }
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Physical Interview", candidateDetails[0][0]);
//      decisionMatrix("physical", row);
      calculateDailySummary("Physical Interview");
      physicalInterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      physicalInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************    
  if(sheetName === "Physical Interview" && column === PHYSICALDECISIONCOLUMN && status === "NO") {
    var result =  ui.alert("Are you sure you want to end the application process for this candidate?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = physicalInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 34).getValues();
      var failedSheetLastRow = getLastRow(failedSheet, 1);
      physicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(failedSheet.getRange(failedSheetLastRow+1, 2));
      failedSheet.getRange(failedSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTATUSCOLUMN).setValue("Failed");
      failedSheet.getRange(failedSheetLastRow+1, FAILEDSTAGECOLUMN).setValue("Physical Interview");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Physical Interview", candidateDetails[0][0]);
      //sendRejectEmail(candidateDetails[0][1], candidateDetails[0][2]);
      calculateDailySummary("Physical Interview");
      physicalInterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      physicalInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************      
  if(sheetName === "Physical Interview" && column === PHYSICALDECISIONCOLUMN && status === "WAITING LIST") {
    var result = ui.alert("Are you sure you want to move this candidate to the waiting list?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = physicalInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 34).getValues();
      var waitingListSheetLastRow = getLastRow(waitingSheet, 1);
      physicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(waitingSheet.getRange(waitingListSheetLastRow+ 1, 2));
      waitingSheet.getRange(waitingListSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
      waitingSheet.getRange(waitingListSheetLastRow+1, WAITINGLISTSTATUSCOLUMN).setValue("Waiting List");
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Physical Interview", candidateDetails[0][0]);
      physicalInterviewSheet.deleteRow(row);
    }
    if (result == ui.Button.NO) {
      physicalInterviewSheet.getRange(row,  column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************        
  if(sheetName === "Physical Interview" && column === PHYSICALDECISIONCOLUMN && status === "MAYBE"){
    var result = ui.alert("Are you sure you want to move this candidate through but flag "+
                          "candidate as a maybe?", ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      var candidateDetails = physicalInterviewSheet.getRange(row, TIMESTAMPCOLUMN, 1, 34).getValues();
      var hiredSheetLastRow = getLastRow(hiredSheet, 1);
      var decision = makeFinalDecision(hiredSheetLastRow+1, candidateDetails[0][21], candidateDetails[0][22], "YES");
      if(decision === "FULL TIME HIRE"){
        physicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(hiredSheet.getRange(hiredSheetLastRow+1, 2));
        hiredSheet.getRange(hiredSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
        hiredSheet.getRange(hiredSheetLastRow+1, 33).setValue("MAYBE");
        hiredSheet.getRange(row, 34).setValue("FULL TIME HIRE");
        hiredSheet.getRange(row, 35).setValue("Awaiting Feedback");
      }
      else if(decision === "CONTRACT"){
        physicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(hiredSheet.getRange(hiredSheetLastRow+ 1, 2));
        hiredSheet.getRange(hiredSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
        hiredSheet.getRange(hiredSheetLastRow+1, 33).setValue("MAYBE");
        hiredSheet.getRange(row, 34).setValue("CONTRACT");
        hiredSheet.getRange(row, 35).setValue("Awaiting Feedback");
      }
      else{
        var vpdecisionSheetLastRow = getLastRow(vpdecisionSheet, 1);
        physicalInterviewSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(vpdecisionSheet.getRange(vpdecisionSheetLastRow+ 1, 2));
        vpdecisionSheet.getRange(vpdecisionSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
        vpdecisionSheet.getRange(vpdecisionSheetLastRow+1, 33).setValue("YES");
        sendVPEMails();
      }
      updateTAT(candidateDetails[0][2], candidateDetails[0][3], "Physical Interview", candidateDetails[0][0]);
//      decisionMatrix("physical", row);
      calculateDailySummary("Physical Interview");
      physicalInterviewSheet.deleteRow(row);
    } 
    if (result == ui.Button.NO) {
      physicalInterviewSheet.getRange(row, column).setValue(e.oldValue);
    } 
  }
//***********************************************************************************************************************************************************************************          
  if(sheetName === "Accepted" && column === 35 && status === "Offer Sent"){
    hiredSheet.getRange(row, 36).setValue("Completed");
  }
//***********************************************************************************************************************************************************************************          
  //Begin from here; Rewrite decision matrix and scheduler for first technical interview and link other stages to failed 
  //and waiting list dcisons from phone and technical interviews
  //Put duplicate checker on every sheet
  if(sheetName === "VP Decision Sheet"){ 
    if(status === "Contract"){
      var result = ui.alert("Are you sure you want to hire this candidate on contract?", ui.ButtonSet.YES_NO);
    
      if (result == ui.Button.YES) {
        var hiredSheetLastRow = getLastRow(hiredSheet, 1);
        vpdecisionSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(hiredSheet.getRange(hiredSheetLastRow+1, 2));
        hiredSheet.getRange(hiredSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
        hiredSheet.getRange(hiredSheetLastRow+1, 33).setValue("VP - YES");
        hiredSheet.getRange(hiredSheetLastRow+1, 34).setValue("CONTRACT");
        hiredSheet.getRange(hiredSheetLastRow+1, 35).setValue("Awaiting Feedback");
        vpdecisionSheet.deleteRow(row);
      } 
      if (result == ui.Button.NO) {
        vpdecisionSheet.getRange(row, column).setValue(e.oldValue);
      } 
    }
    if(status === "Full-Time Hire"){
      var result = ui.alert("Are you sure you want to hire this candidate full time?", ui.ButtonSet.YES_NO);
    
      if (result == ui.Button.YES) {
        var hiredSheetLastRow = getLastRow(hiredSheet, 1);
        vpdecisionSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(hiredSheet.getRange(hiredSheetLastRow+1, 2));
        hiredSheet.getRange(hiredSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
        hiredSheet.getRange(hiredSheetLastRow+1, 33).setValue("VP - YES");
        hiredSheet.getRange(hiredSheetLastRow+1, 34).setValue("FULL TIME HIRE");
        hiredSheet.getRange(hiredSheetLastRow+1, 35).setValue("Awaiting Feedback");
        vpdecisionSheet.deleteRow(row);
      } 
      if (result == ui.Button.NO) {
        vpdecisionSheet.getRange(row, column).setValue(e.oldValue);
      } 
    }
    if(status === "Reject"){
      var result = ui.alert("Are you sure you want to reject this candidate?", ui.ButtonSet.YES_NO);
    
      if (result == ui.Button.YES) {
        var failedSheetLastRow = getLastRow(failedSheet, 1);
        failedSheet.getRange(failedSheetLastRow+1, TIMESTAMPCOLUMN).setValue(new Date());
        vpdecisionSheet.getRange(row, EMAILADDRESSCOLUMN, 1, 31).copyTo(failedSheet.getRange(failedSheetLastRow+1, 2));
        failedSheet.getRange(failedSheetLastRow+1, FAILEDSTATUSCOLUMN).setValue("Failed");
        failedSheet.getRange(failedSheetLastRow+1, FAILEDSTAGECOLUMN).setValue("VP Decision");
        //sendRejectEmail(candidateDetails[0][0], candidateDetails[0][1]);
        vpdecisionSheet.deleteRow(row);
      } 
      if (result == ui.Button.NO) {
        vpdecisionSheet.getRange(row, column).setValue(e.oldValue);
      } 
    }
  }
}
//***********************************************************************************************************************************************************************************          
function updateTAT(name, role, stage, timeOfEntry){
  var timeOfExit = new Date();
  var lastRow = getLastRow(tatSheet, 1);
  var currentRow = 2;
  var values = tatSheet.getRange(currentRow, 1, lastRow, 2).getValues();
  
  for(var i = 0; i <= values.length-1; i++){
    if(values[i][0] === name && values[i][1] === role){
      var duration = Math.round((timeOfExit.getTime() - timeOfEntry.getTime()) / (1000*60*60*24));
      if(stage === "Prospectives"){
        tatSheet.getRange(i+2, 3).setValue(duration);
      }
      if(stage === "Portfolio Review"){
        tatSheet.getRange(i+2, 4).setValue(duration);
      }
      if(stage === "Phone Interview"){
        tatSheet.getRange(i+2, 5).setValue(duration);
      }
      if(stage === "First Technical Interview"){
        tatSheet.getRange(i+2, 6).setValue(duration);
      }
      if(stage === "Technical Interview"){
        tatSheet.getRange(i+2, 7).setValue(duration);
      }
      if(stage === "Physical Interview"){
        tatSheet.getRange(i+2, 8).setValue(duration);
      }
      return;
    }
  }
  
  tatSheet.getRange(lastRow+1, 1).setValue(name);
  tatSheet.getRange(lastRow+1, 2).setValue(role);
  var duration = Math.round((timeOfExit.getTime() - timeOfEntry.getTime()) / (1000*60*60*24));
  if(stage === "Prospectives"){
    tatSheet.getRange(lastRow+1, 3).setValue(duration);
  }
  if(stage === "Porfolio Review"){
    tatSheet.getRange(lastRow+1, 4).setValue(duration);
  }
  if(stage === "Phone Interview"){
    tatSheet.getRange(lastRow+1, 5).setValue(duration);
  }
   if(stage === "First Technical Interview"){
    tatSheet.getRange(lastRow+1, 6).setValue(duration);
  }
  if(stage === "Technical Interview"){
    tatSheet.getRange(lastRow+1, 7).setValue(duration);
  }
  if(stage === "Physical Interview"){
    tatSheet.getRange(lastRow+1, 8).setValue(duration);
  }
}
//***********************************************************************************************************************************************************************************          
function makeFinalDecision(row, phone, skype, physical){
  if(skype !== ""){
    if((phone === "YES" && skype === "YES" && physical === "YES") || (phone === "MAYBE" && skype === "YES" && physical === "YES")){
      return "FULL TIME HIRE";
    }
    if((phone === "YES" && skype === "YES" && physical === "MAYBE") || (phone === "YES" && skype === "MAYBE" && physical === "MAYBE")
      || (phone === "MAYBE" && skype === "YES" && physical === "MAYBE") || (phone === "MAYBE" && skype === "MAYBE" && physical === "YES")
      || (phone === "MAYBE" && skype === "MAYBE" && physical === "MAYBE")){
      return "VP TO DECIDE";
    }
    if(phone === "YES" && skype === "MAYBE" && physical === "YES"){
      return "CONTRACT";
    }
  }
  if(skype === ""){
    if(phone === "YES" && physical === "YES"){
      return "FULL TIME HIRE";
    }
    if((phone === "YES" && physical === "MAYBE") || (phone === "MAYBE" && physical === "YES") || (phone === "MAYBE" && physical === "MAYBE")){
      return "VP TO DECIDE";
    }
  }
}
//***********************************************************************************************************************************************************************************          
function calculateDailySummary(stage){
  var lastRow = getLastRow(summarySheet, 1);
  var values = summarySheet.getRange(2, 1, lastRow, 1).getValues();
  var date = new Date();
  var day = date.getDate();
  var month = date.getMonth()+1;
  var year = date.getFullYear();
  
  for(var i = 0; i <= values.length-1; i++){
    var temp = new Date(values[i][0]);
    if((temp.getDate() === day) && (temp.getMonth()+1 === month) && (temp.getFullYear() === year)){
      if(stage === "Phone Interview"){
        var oldValue = summarySheet.getRange(i+2, 2).getValue();
        summarySheet.getRange(i+2, 2).setValue(oldValue+1);
      }
      if(stage === "Porfolio Review"){
        var oldValue = summarySheet.getRange(i+2, 3).getValue();
        summarySheet.getRange(i+2, 3).setValue(oldValue+1);
      }
      if(stage === "First Technical Interview"){
        var oldValue = summarySheet.getRange(i+2, 4).getValue();
        summarySheet.getRange(i+2, 4).setValue(oldValue+1);
      }
      if(stage === "Technical Interview"){
        var oldValue = summarySheet.getRange(i+2, 5).getValue();
        summarySheet.getRange(i+2, 5).setValue(oldValue+1);
      }
      if(stage === "Physical Interview"){
        var oldValue = summarySheet.getRange(i+2, 6).getValue();
        summarySheet.getRange(i+2, 6).setValue(oldValue+1);
      }
      return;
    }
  }
  summarySheet.getRange(lastRow+1, 1).setValue(new Date(month+"/"+day+"/"+year));
  if(stage === "Phone Interview"){
    var oldValue = summarySheet.getRange(lastRow+1, 2).getValue();
    summarySheet.getRange(lastRow+1, 2).setValue(oldValue+1);
  }
  if(stage === "Portfolio Review"){
    var oldValue = summarySheet.getRange(lastRow+1, 3).getValue();
    summarySheet.getRange(lastRow+1, 3).setValue(oldValue+1);
  }
  if(stage === "First Technical Interview"){
    var oldValue = summarySheet.getRange(lastRow+1, 4).getValue();
    summarySheet.getRange(lastRow+1, 4).setValue(oldValue+1);
  }
  if(stage === "Technical Interview"){
    var oldValue = summarySheet.getRange(lastRow+1, 5).getValue();
    summarySheet.getRange(lastRow+1, 5).setValue(oldValue+1);
  }
  if(stage === "Physical Interview"){
    var oldValue = summarySheet.getRange(lastRow+1, 6).getValue();
    summarySheet.getRange(lastRow+1, 6).setValue(oldValue+1);
  }
}
//***********************************************************************************************************************************************************************************          
function getLastRow(sheet, colToCheck){
  var lastRow = sheet.getMaxRows();
  var values = sheet.getRange(1, colToCheck, lastRow).getValues();
  
  while(values[lastRow - 1] == "" && lastRow > 0){
    lastRow--;
  }
  
  return lastRow;
}
//***********************************************************************************************************************************************************************************          
function sendRejectEmail_Phone(name, recipient){
  var subject = "APPLICATION UPDATE";
  var body = "Dear "+name+",\n\nWe appreciate the interest in joining our journey to financial happiness. Please note that we consider factors around clarity,"
  +"depth and experience (in relevance to the role applied for) in shortlisting candidates.\n\nFollowing your recent phone interview, we will not be moving ahead "
  +"with your application at this time due to a lack in one or more of the factors listed above.\n\nWe however encourage you to continue to build yourself and "
  +"we appreciate the time you have invested in your application process with TeamApt.\n\nBest wishes in your future endeavours.\n\nBest,\nRecruitment Team,\nTeamApt.";
  MailApp.sendEmail(recipient, subject, body, {
    name: "TeamApt Recruitment Bot",
    noReply: true
  });
}
//***********************************************************************************************************************************************************************************          
function sendRejectEmail_Technical(name, recipient){
  var subject = "APPLICATION UPDATE";
  var body = "Dear "+name+",\n\nThank you very much for the interest in joining our journey to financial happiness and we congratulate you for coming this far "
  +"in your application process.\n\nFollowing your recent Technical interview, we are afraid your performance does not measure up against our requirements, and thus, "
  +"we cannot proceed to the final stage of the process.\n\nWe encourage you to continue to build yourself and we appreciate the time you have invested in your "
  +"application process with TeamApt.\n\nBest wishes in your future endeavours.\n\nBest Regards,\nRecruitment Team,\nTeamApt.";
  MailApp.sendEmail(recipient, subject, body, {
    name: "TeamApt Recruitment Bot",
    noReply: true
  });
}
//***********************************************************************************************************************************************************************************          
function moveToTechnicalInterviewEmail(name, recipient){
  var subject = "APPLICATION UPDATE";
  var body = "Dear "+name+",\n\nWe appreciate the time you have invested so far in the recruitment process and we are excited "
  +"to inform you that you have been shortlisted to move to the next stage, which is a technical interview. Here your skill "
  +"level will be tested with a simple challenge.\n\nKindly check your email frequently as more information will be communicated "
  +"to you in due time.\n\nWe wish you best of luck.\n\nBest Regards,\nRecruitment Team,\nTeamApt.";
  MailApp.sendEmail(recipient, subject, body, {
    name: "TeamApt Recruitment Bot",
    noReply: true
  });
}
//***********************************************************************************************************************************************************************************          
function moveToPhysicalInterviewEmail(name, recipient){
  var subject = "APPLICATION UPDATE";
  var body = "Dear "+name+",\n\nWe appreciate the time you have invested so far in the recruitment process and we are excited "
  +"to inform you that you have been shortlisted to move to the next stage, which is a physical Interview with the MD.\n\nA "
  +"member of our team will reach out to you to agree on a suitable time for the interview. Kindly check your email frequently "
  +"as more information will be communicated to you in due time.\n\nWe wish you best of luck.\n\nBest Regards,\nRecruitment Team,\nTeamApt.";
  MailApp.sendEmail(recipient, subject, body, {
    name: "TeamApt Recruitment Bot",
    noReply: true
  });
}
//***********************************************************************************************************************************************************************************          
function sendNotificationEmail(type){
  if(type == "internship"){
    var sheetlink = "https://docs.google.com/spreadsheets/d/1VUXFWVh27QlaZoE9uhm18Ffsg_D4AK3yLWeyV5eHJTY/edit#gid=1828302115";
  }
  
  if(type == "scholarship"){
    var sheetlink = "https://docs.google.com/spreadsheets/d/1VUXFWVh27QlaZoE9uhm18Ffsg_D4AK3yLWeyV5eHJTY/edit#gid=1621046468";
  }
    
  var recipient = "hr@teamapt.com";
  var subject = "NEW APPLICATION ENTRY";
  var body = "The internship/scholarship application form has been filled, go to "+sheetlink+" for more details";
  
  MailApp.sendEmail(recipient, subject, body, {
    name: "TeamApt Notification Bot",
    noReply: true
  });
}
//***********************************************************************************************************************************************************************************          
function sendReceiptEmail(name, recipient){
  var subject = "APPLICATION RECEIVED";
  var body = "Hi "+name+",\n\nYou are receiving this email because you just submitted a job application form to TeamApt Limited.\n\n"+
    "This is to let you know that your application has been received and is currently under review.\n\n"+ "Kindly note that if you are shortlisted you will be contacted soon.\n\n"+
      "Thanks for your interest in TeamApt and goodluck!\n\nRegards,\nRecruitment Team.";
  MailApp.sendEmail(recipient, subject, body, {
    name: "TeamApt Recruitment Bot",
    noReply: true
  });
}
//***********************************************************************************************************************************************************************************          
function sendVPEMails(){
  var sheetLink = "https://docs.google.com/spreadsheets/d/1VUXFWVh27QlaZoE9uhm18Ffsg_D4AK3yLWeyV5eHJTY/edit#gid=888934852";
  var subject = "URGENT DECISION TO BE MADE";
  var vps = ["Emeka Ibe", "Simpa Saiki"];
  for(var i = 0; i <= vps.length-1; i++){
    var body = "Hello "+vps[i]+",\n\nA new entry is waiting for a decision to be made by you under the “VP Decision Sheet” in the Hiring workbook.\n\n"+
    "As part of our recruitment process and based on our decision matrix, applications may be forwarded to you for review. You are expected to  make a "+
    "final decision on those applications. Kindly access the link here "+sheetLink+"\n\nRegards,\nRecruitment Team";
    MailApp.sendEmail(getEmployeeEmail(vps[i]), subject, body, {
      name : "TeamApt Notification Bot"
    });
  }
}
//***********************************************************************************************************************************************************************************          
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
//***********************************************************************************************************************************************************************************          
function decisionMatrix(sheet, row){
  var candidateValues, interviewerValues, physicalInterviewDecision;
  var lastRow = getLastRow(decisionMatrixSheet, 1);
  var emailAddresses = decisionMatrixSheet.getRange(2, 1, lastRow, 1).getValues()
  
  switch(sheet){
    case "phone": candidateValues = phoneInterviewSheet.getRange(row, 2, 1, 3).getValues();
      interviewerValues = phoneInterviewSheet.getRange(row, 10, 1, 4).getValues();
      break;
    case "technical": candidateValues = technicalInterviewSheet.getRange(row, 2, 1, 3).getValues();
      interviewerValues = technicalInterviewSheet.getRange(row, 14, 1, 4).getValues();
      break;
    case "physical": candidateValues = physicalInterviewSheet.getRange(row, 2, 1, 3).getValues();
      interviewerValues = physicalInterviewSheet.getRange(row, 18, 1, 7).getValues();
      physicalInterviewDecision = physicalInterviewSheet.getRange(row, 27).getValue();
      break;
  }
  
  for(var i = 0; i <= emailAddresses.length-1; i++){
    if(emailAddresses[i][0] === candidateValues[0][0] && sheet === "phone"){
      decisionMatrixSheet.getRange(i+2, 1, 1, 3).setValues(candidateValues);
      decisionMatrixSheet.getRange(i+2, 4, 1, 4).setValues(interviewerValues);
      return;
    }
    if(emailAddresses[i][0] === candidateValues[0][0] && sheet === "technical"){
      decisionMatrixSheet.getRange(i+2, 1, 1, 3).setValues(candidateValues);
      decisionMatrixSheet.getRange(i+2, 8, 1, 4).setValues(interviewerValues);
      return;
    }
    if(emailAddresses[i][0] === candidateValues[0][0] && sheet === "physical"){
      decisionMatrixSheet.getRange(i+2, 1, 1, 3).setValues(candidateValues);
      decisionMatrixSheet.getRange(i+2, 12, 1, 7).setValues(interviewerValues);
      decisionMatrixSheet.getRange(i+2, 19).setValue(physicalInterviewDecision);
      return;
    }
  }
  
  if(sheet === "phone"){
    decisionMatrixSheet.getRange(lastRow+1, 1, 1, 3).setValues(candidateValues);
    decisionMatrixSheet.getRange(lastRow+1, 4, 1, 4).setValues(interviewerValues);
  }
  if(sheet === "technical"){
    decisionMatrixSheet.getRange(lastRow+1, 1, 1, 3).setValues(candidateValues);
    decisionMatrixSheet.getRange(lastRow+1, 8, 1, 4).setValues(interviewerValues);
  }
  if(sheet === "physical"){
    decisionMatrixSheet.getRange(lastRow+1, 1, 1, 3).setValues(candidateValues);
    decisionMatrixSheet.getRange(lastRow+1, 12, 1, 7).setValues(interviewerValues);
    decisionMatrixSheet.getRange(lastRow+1, 19).setValue(physicalInterviewDecision);
  }
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------  

function isDuplicate(name, sheet, colToCheck){
  var lastRow = getLastRow(sheet, colToCheck);
  var names = sheet.getRange(2, 2, lastRow, 1).getValues();

  for(var i = 2; i <= lastRow; i++){
    if(names[i-2][0] === name){
      return i;
    }
  }
  return "Not duplicate";
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------  

function scheduler(type, candidateFunction, candidateEmail, candidatePhoneNumber, candidateCV, candidatePortfolio){
  if(type === "Physical"){
    physicalScheduler(candidateFunction, candidateEmail, candidatePhoneNumber, candidateCV);
    return;
  }
  var schedule = {};
  var lowest, atRow, interviewer, callTime;
  var interviewerFunctions, interviewsTaken, interviewerTypes, interviewersLastCallTime;
  switch(type){
    case "Phone": var lastRow = getLastRow(interviewerSheet, INTERVIEWERNAMECOLUMN);
      interviewerFunctions = interviewerSheet.getRange(2, INTERVIEWERFUNCTIONSCOLUMN, lastRow, 1).getValues();
      interviewersLastCallTime = interviewerSheet.getRange(2, INTERVIEWERCALLTIMECOLUMN, lastRow, 1).getValues();
      interviewsTaken = interviewerSheet.getRange(2, INTERVIEWERCALLSTAKENCOLUMN, lastRow, 1).getValues();
      break;
    case "Portfolio Review": var lastRow = getLastRow(interviewerSheet, INTERVIEWERNAMECOLUMN);
      interviewerFunctions = interviewerSheet.getRange(2, INTERVIEWERFUNCTIONSCOLUMN, lastRow, 1).getValues();
      interviewersLastCallTime = interviewerSheet.getRange(2, INTERVIEWERCALLTIMECOLUMN, lastRow, 1).getValues();
      interviewsTaken = interviewerSheet.getRange(2, INTERVIEWERCALLSTAKENCOLUMN, lastRow, 1).getValues();
      break;
    case "First Technical": var lastRow = getLastRow(technicalInterviewerSheet, INTERVIEWERNAMECOLUMN);
      interviewerFunctions = technicalInterviewerSheet.getRange(2, INTERVIEWERFUNCTIONSCOLUMN, lastRow, 1).getValues();
      interviewersLastCallTime = technicalInterviewerSheet.getRange(2, INTERVIEWERCALLTIMECOLUMN, lastRow, 1).getValues();
      interviewsTaken = technicalInterviewerSheet.getRange(2, INTERVIEWERCALLSTAKENCOLUMN, lastRow, 1).getValues();
      break;
    case "Technical": var lastRow = getLastRow(technicalInterviewerSheet, INTERVIEWERNAMECOLUMN);
      interviewerFunctions = technicalInterviewerSheet.getRange(2, INTERVIEWERFUNCTIONSCOLUMN, lastRow, 1).getValues();
      interviewersLastCallTime = technicalInterviewerSheet.getRange(2, INTERVIEWERCALLTIMECOLUMN, lastRow, 1).getValues();
      interviewsTaken = technicalInterviewerSheet.getRange(2, INTERVIEWERCALLSTAKENCOLUMN, lastRow, 1).getValues();
      break;
  }
  
  for(var i = 2; i <= lastRow; i++){
    var pattern = candidateFunction;
    var regEx = new RegExp(pattern);
    if(regEx.test(interviewerFunctions[i-2][0])){
      if(lowest === undefined && atRow === undefined){
        lowest = interviewsTaken[i-2][0];
        atRow = i;
      }
      else{
        if(interviewsTaken[i-2][0] < lowest){
          lowest = interviewsTaken[i-2][0];
          atRow = i;
        }
      }
    }
  }
  switch(type){
    case "Phone": interviewer = interviewerSheet.getRange(atRow, INTERVIEWERNAMECOLUMN).getValue();
      break;
    case "Portfolio Review": interviewer = interviewerSheet.getRange(atRow, INTERVIEWERNAMECOLUMN).getValue();
      break;
    case "First Technical": interviewer = technicalInterviewerSheet.getRange(atRow, INTERVIEWERNAMECOLUMN).getValue();
      break;
    case "Technical": interviewer = technicalInterviewerSheet.getRange(atRow, INTERVIEWERNAMECOLUMN).getValue();
      break;
  }
  
  callTime = new Date(interviewersLastCallTime[atRow-2][0]);
  var hour = callTime.getHours();
  
  if(hour === 16){
    if(callTime.getDay()+1 === 6){
      callTime.setDate(callTime.getDate()+3);
      callTime.setHours(15,0,0);
    }
    else{
      callTime.setDate(callTime.getDate()+1);
      callTime.setHours(15,0,0);
    }
  }
  if(hour === 15){
    callTime.setHours(16, 0, 0);
  }
  
  schedule.interviewer = interviewer;
  schedule.callTime = callTime;
  sendInvites(type, candidateEmail, candidatePhoneNumber, candidateCV, interviewer, callTime, candidatePortfolio);
  updateSummarySheet(interviewer,atRow, type, callTime);
  return schedule;
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------  

function updateSummarySheet(interviewer, atRow, type, callTime){
  var lastRow = getLastRow(summarySheet, 1);
  var names = summarySheet.getRange(2, 1, lastRow, 1).getValues();
  switch(type){
    case "Phone": interviewerSheet.getRange(atRow, INTERVIEWERCALLSTAKENCOLUMN).setValue(interviewerSheet.getRange(atRow, INTERVIEWERCALLSTAKENCOLUMN).getValue()+1);
      break;
    case "Portfolio Review": interviewerSheet.getRange(atRow, INTERVIEWERCALLSTAKENCOLUMN).setValue(interviewerSheet.getRange(atRow, INTERVIEWERCALLSTAKENCOLUMN).getValue()+1);
      break;
    case "Technical": technicalInterviewerSheet.getRange(atRow, INTERVIEWERCALLSTAKENCOLUMN).setValue(technicalInterviewerSheet.getRange(atRow, INTERVIEWERCALLSTAKENCOLUMN).getValue()+1);
      break;
    case "Physical": physicalInterviewerSheet.getRange(atRow, INTERVIEWERCALLSTAKENCOLUMN).setValue(physicalInterviewerSheet.getRange(atRow, INTERVIEWERCALLSTAKENCOLUMN).getValue()+1);
      break;
  }
  for(var i = 2; i <= lastRow; i++){
    if(names[i-2][0] === interviewer){
      summarySheet.getRange(i, 2).setValue(callTime);
      return;
    }
  }
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------  

function isWeekday() {
  var today = new Date();
  var day = today.getDay();
  if (today == 6 || today == 0) {
    return false;
  }
  else {
    return true;
  }
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------  

function sendInvites(type, candidateEmail, candidatePhoneNumber, candidateCV, interviewer, callTime, candidatePortfolio){
  var interviewerEmail = getEmployeeEmail(interviewer);
  var endTime = new Date(callTime);
  var hour = endTime.getHours();
  endTime.setHours(hour, 30, 0);
  var guestEmails = candidateEmail + "," + interviewerEmail;
  
  if(type === "Phone"){
    CalendarApp.getCalendarById("recruitment@teamapt.com").createEvent('TeamApt Limited - Phone Interview', callTime, endTime, {
      description: 'CV: ' + candidateCV + '\n\nPhone: ' + candidatePhoneNumber,
      guests: guestEmails,
      sendInvites: true
    });
  }
  if(type === "Portfolio Review"){
    CalendarApp.getCalendarById("recruitment@teamapt.com").createEvent('TeamApt Limited - Portfolio Review', callTime, endTime, {
      description: 'CV: ' + candidateCV + '\n\nPortfolio: ' + candidatePortfolio,
      guests: guestEmails,
      sendInvites: true
    });
  }
  if(type === "Technical"){
    CalendarApp.getCalendarById("recruitment@teamapt.com").createEvent('TeamApt Limited - Technical Interview', callTime, endTime, {
      description: 'CV: ' + candidateCV + '\n\nPhone: ' + candidatePhoneNumber,
      guests: guestEmails,
      sendInvites: true
    });
  }
}
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------  
function physicalScheduler(candidateFunction, candidateEmail, candidatePhoneNumber, candidateCV){
  var date;
  var counter = physicalInterviewerSheet.getRange(7, 4).getValue();
  var guestEmails = candidateEmail+",ooguntola@teamapt.com, cachugbu@teamapt.com";
  if(summarySheet.getRange(19, 2).getValue() === ""){
    date = new Date();
    date.setDate(date.getDate()+2);
    switch(date.getDay()){
      case 0: date.setDate(date.getDate()+1); 
        break;
      case 6: date.setDate(date.getDate()+2);
        break;
    }
  }
  else{
    date = new Date(summarySheet.getRange(19, 2).getValue());
  }
  
  if(counter === 2){
    physicalInterviewerSheet.getRange(7, 4).setValue(0);
    date.setDate(date.getDate()+1); 
    switch(date.getDay()){
      case 0: date.setDate(date.getDate()+1); 
        break;
      case 6: date.setDate(date.getDate()+2);
        break;
    }
  }
  
  while(!isTosinFree(date)){
    date.setDate(date.getDate()+1);
    switch(date.getDay()){
      case 0: date.setDate(date.getDate()+1); 
        break;
      case 6: date.setDate(date.getDate()+2);
        break;
    }
  }
  
  var timeOfInterview = getTimeOfInterview(date);
  var endTime = new Date(timeOfInterview);
  var hour = endTime.getHours();
  endTime.setHours(hour, 30, 0);
  CalendarApp.getCalendarById("recruitment@teamapt.com").createEvent('TeamApt Limited - Physical Interview', timeOfInterview, endTime, {
    description: 'CV: ' + candidateCV + '\n\nPhone: ' + candidatePhoneNumber,
    guests: guestEmails,
    sendInvites: true
  });
  
  summarySheet.getRange(19, 2).setValue(new Date(timeOfInterview));
  counter = physicalInterviewerSheet.getRange(7, 4).getValue();
  physicalInterviewerSheet.getRange(7, 4).setValue(counter+1);
}

function isTosinFree(date){
  var timeOfInterview;
  var takenTimes = [];
  var calendar = CalendarApp.getCalendarById("teniolorunda@teamapt.com");
  var startTime = new Date(date); var endTime = new Date(date);
  startTime.setHours(8,0,0);
  endTime.setHours(14,0,0);
  var eventsForToday = calendar.getEvents(startTime, endTime);
  for(i in eventsForToday){
    takenTimes.push(eventsForToday[i].getStartTime().getHours());
    takenTimes.push(eventsForToday[i].getEndTime().getHours());
  }
  for(var j = 8; j <= 14; j++){
    if(!isPresent(j, takenTimes)){
      timeOfInterview = j;
      break;
    }
  }
  
  if(timeOfInterview === "" || timeOfInterview === undefined){
    return false;
  }
  else{
    return true;
  }
}

function getTimeOfInterview(date){
  var timeOfInterview;
  var takenTimes = [];
  var calendar = CalendarApp.getCalendarById("teniolorunda@teamapt.com");
  var startTime = new Date(date); var endTime = new Date(date);
  startTime.setHours(8,0,0);
  endTime.setHours(14,0,0);
  var eventsForToday = calendar.getEvents(startTime, endTime);
  for(i in eventsForToday){
    takenTimes.push(eventsForToday[i].getStartTime().getHours());
    takenTimes.push(eventsForToday[i].getEndTime().getHours());
  }
  for(var j = 8; j <= 14; j++){
    if(!isPresent(j, takenTimes)){
      timeOfInterview = j;
      break;
    }
  }
  date.setHours(timeOfInterview, 0, 0);
  return date;
}

function isPresent(number, array){
  for(var i = 0; i <= array.length-1; i++){
    if(number === array[i]){
      return true;
    }
  }
  return false;
}