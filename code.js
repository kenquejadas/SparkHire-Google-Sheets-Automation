function main() {
  const formSpreadsheet = SpreadsheetApp.openById(" Form Spreadsheet ID here ").getActiveSheet();
  const lastRow = formSpreadsheet.getLastRow();
  var mainSpreadsheet;    //stores the hm spreadsheet returned from createSpreadsheet()
  var numReviewers = 0;   
  var reviewerEmails = [];    
  var reviewerNames = [];    
  var questions = [];
  var numQuestions;   //the number of interview questions
  const positionTitle = formSpreadsheet.getRange(lastRow, 3).getValue();    
  const formUserEmail = formSpreadsheet.getRange(lastRow, 2).getValue();    
  const hiringManagerEmail = formSpreadsheet.getRange(lastRow, 5).getValue();    
  const hiringManagerName = formSpreadsheet.getRange(lastRow, 4).getValue();   
  var candidateListSpreadsheet;   
  var formUserMessage;    //stores the message to be emailed to the form user


  const lock = LockService.getScriptLock();
  try {
    while(lock.hasLock() === false){
      lock.waitLock(45000);
    }
  }
  catch (e) {
    Logger.log('Could not obtain lock after 45 seconds.');
   return -1;
  }


  let curCell = formSpreadsheet.getRange(1, 12);
  while(curCell.getValue() != ""){
      numReviewers++;
      curCell = formSpreadsheet.getRange(1, 12 + numReviewers);
  }
  for(let x = 0; x < numReviewers; x+=2){
    curCell = formSpreadsheet.getRange(lastRow, 12+x);
    if(curCell.getValue() != ""){
      reviewerNames.push(curCell.getValue());
    }
  }
  for(let x = 1; x < numReviewers; x+=2){
    curCell = formSpreadsheet.getRange(lastRow, 12+x);
    if(curCell.getValue() != ""){
      reviewerEmails.push(curCell.getValue());
    }
  }
  numReviewers = reviewerEmails.length;

  if(reviewerEmails.length != reviewerNames.length){
    MailApp.sendEmail({
    to: formUserEmail,
    subject: "Error with Submission of the Spark Hire Form",
    htmlBody: "When parsing the responses of the form submission, we found that the number of reviewer emails given was not equal to the number of reviewer names. Please try submitting the form again and ensuring the fields are filled correctly."
    })
    lock.releaseLock();
    return 0;
  }

  for(let x = 0; x < 6; x++){
    curCell = formSpreadsheet.getRange(lastRow, 6+x);
    if(curCell.getValue() != ""){
      questions.push(curCell.getValue());
    }
  }
  numQuestions = questions.length;


  candidateListSpreadsheet = SpreadsheetApp.create("Spark Hire Review Candidate List - " + positionTitle);
  let candSheet = candidateListSpreadsheet.getSheetByName("Sheet1");
  candSheet.getRange("A1").setValue('First & Last Name').setFontWeight("bold");
  candSheet.setColumnWidth(1, 200);

  mainSpreadsheet = createSpreadsheet(positionTitle, questions, numQuestions, numReviewers, candidateListSpreadsheet.getUrl(), hiringManagerName);
  
  formUserMessage = createShareLinkReviewerSpreadsheets(reviewerNames, reviewerEmails, mainSpreadsheet, numQuestions, positionTitle, candidateListSpreadsheet, formUserEmail, hiringManagerEmail);

  formUserMessage += addEmailHm(hiringManagerEmail, formUserEmail, mainSpreadsheet, candidateListSpreadsheet, positionTitle);


  //email form user with results of running the script
  MailApp.sendEmail({
    to: formUserEmail,
    subject: "Results of Submitting the Spark Hire Form",
    htmlBody: "Thank you for filling out the Spark Hire Form. Below are the results. If any errors occurred and required you to manually add the hiring manager or a reviewer to any spreadsheets, please remove yourself from those spreadsheets after doing so. <br><br>" + formUserMessage + "<br><br>Here is a link to the candidate list. You will only have access to this if an issue occurred when sharing the spreadsheets with the hiring manager or reviewers. Have a nice day!<br><br><a href=\"" + candidateListSpreadsheet.getUrl() + "\">Candidate List Spreadsheet</a>"
  })

  lock.releaseLock();
  return 0;

}




function createSpreadsheet(positionTitle, questions, numQuestions, numReviewers, candListUrl, hiringManagerName){
  //Creates the hiring manager's spreadsheet and returns it
  //Links sheet1 and combined sheet to the candidate list
  //getRange(row, column, numRows, numColumns)

  var mainSpreadsheet = SpreadsheetApp.create(hiringManagerName + "'s Spark Hire Interview Review Sheet - " + positionTitle);
  var sheet1 = mainSpreadsheet.getSheetByName("Sheet1");

  //Add in default labels and make cells bolded
  sheet1.getRange('A1').setValue("Below Expectations (1) \nMeets Expectations (2) \nExceeds Expectations (3)");
  sheet1.getRange('A2').setValue("First & Last Name");
  sheet1.getRange(2, numQuestions + 2, 1, 1).setValue("Total points\nout of " + (numQuestions * 3));
  sheet1.getRange(2, numQuestions + 3, 1, 1).setValue("Notes");
  sheet1.getRange(1, 1, 2, numQuestions + 3).setFontWeight("bold");

  //Add questions
  for (let i = 0; i < numQuestions; i++) {
      sheet1.getRange(2, 2 + i).setValue(questions[i]);
  }

  //Add in default highlighting
  sheet1.getRange(1, 1, 1, numQuestions + 3).setBackground("#d9d9d9");
  sheet1.getRange(2, 1, 1, numQuestions + 1).setBackground("#fff2cc");
  sheet1.getRange(2, numQuestions + 2, 1, 2).setBackground("#ffe599");

  //Adjust column widths and row height
  sheet1.setColumnWidths(1, numQuestions + 1, 255);
  sheet1.setColumnWidth(numQuestions + 2, 85);
  sheet1.setColumnWidth(numQuestions + 3, 275);
  sheet1.setRowHeight(2, 125);
  sheet1.getRange(2, 2, 1, numQuestions).setWrap(true);

  //link name column to candidate list
  sheet1.getRange(3,1).setFormula('=IMPORTRANGE("' + candListUrl + '","A2:A1000")');

  //protect candidate names column
  var protect1 = sheet1.getRange("A:A").protect().setDescription("Candidate name column can't be edited to ensure proper synchronization of the hiring manager's combined results sheet.");
  protect1.removeEditors(protect1.getEditors());


  //setup sum formula
  let formulaRange = "B3";
  let curchar = 66 + numQuestions - 1;
  formulaRange += ":" + String.fromCharCode(curchar) + "3";
  sheet1.getRange(3,numQuestions + 2).setFormula('=SUM(' + formulaRange + ')');
  sheet1.getRange(3,numQuestions + 3).setValue("<-- Click on this cell, then click on the circle in the bottom right and drag it down to apply the sum function to more rows. In case you accidentally delete the formula, it is =SUM(" + formulaRange + ")");

  //Create and setup combined results sheet
  var combinedSheet = mainSpreadsheet.insertSheet("Combined Results");
  combinedSheet.getRange(1, 1, 2, 1).setBackground('#fff2cc').merge().setValue("First & Last Name");
  combinedSheet.getRange(1, 2, 2, numReviewers+1).setBackground('#ffe599');
  combinedSheet.getRange(1, numReviewers+2, 1, 1).setValue("Yours").setHorizontalAlignment("center");
  combinedSheet.getRange(1, numReviewers+3, 2, 1).setBackground('#ffd966').merge().setValue("Average Score\nout of " + (numQuestions * 3)).setHorizontalAlignment("center");
  combinedSheet.getRange(1, 2, 1, numReviewers).merge().setValue("Reviewers").setHorizontalAlignment("center");
  combinedSheet.getRange(1, 1, 2, numReviewers+3).setFontWeight("bold");
  combinedSheet.setColumnWidth(1, 255);
  combinedSheet.setColumnWidths(2, numReviewers, 150);
  combinedSheet.getRange(2,numReviewers + 2).setValue(hiringManagerName);
  combinedSheet.getRange(3,1).setFormula('=IMPORTRANGE("' + candListUrl + '","A2:A1000")');

  //setup average formula
  formulaRange = "B3";
  curchar = 66 + numReviewers;
  formulaRange += ":" + String.fromCharCode(curchar) + "3";
  combinedSheet.getRange(3,numReviewers + 3).setFormula('=AVERAGEIF(' + formulaRange + ',"<>0")');
  combinedSheet.getRange(3,numReviewers+4).setValue("<-- Click on this cell, then click on the circle in the bottom right and drag it down to apply the average function to more rows. In case you accidentally delete the formula, it is =AVERAGEIF(" + formulaRange + ",\"<>0\")");

  //protect combined results sheet
  var protect2 = combinedSheet.getRange(3,1,997,numReviewers+2).protect().setDescription("Modifications aren't allowed to help with error prevention.");
  protect2.removeEditors(protect2.getEditors());
  

  return mainSpreadsheet;
}




function createShareLinkReviewerSpreadsheets(reviewerNames, reviewerEmails, mainSpreadsheet, numQuestions, positionTitle, candidateList, formUserEmail, hiringManagerEmail){
  //This function creates a spreadsheet for each reviewer and links each spreadsheet to the combined results sheet belonging to the hiring manager. Each reviewer is sent an email as well informing them that they have been included as a reviewer. Other minor things performed in the function: add reviewer names to combined sheet, link hm scores to combined sheet

  //The function returns a string to be included in the email to the form user. If there was an error with sharing a spreadsheet or sending an email, this message includes the name, email, and link to a spreadsheet for each reviewer that there was an error for so that the form user can manually share the form with the reviewer.

  var reviewerEmailSubject = "You have been added as a reviewer for the \"" + positionTitle + "\" position.";
  var reviewerEmailBody = "Hello,<br><br> You are receiving this email to inform you that you have been included as a reviewer in the Spark Hire hiring process for the \"" + positionTitle + "\" position. You should have been added as an editor of a spreadsheet where you can score applicants' responses to interview questions. This spreadsheet will be automatically populated with candidate first and last names as they are added to the canidate list spreadsheet for the position. Your spreadsheet's \"Total Points\" column is linked to a spreadsheet belonging to the hiring manager for the position so that their spreadsheet will automatically be populated with the total score you give to each applicant. Attached below are the links to your spreadsheet and the candidate list spreadsheet. Do not make any modifications to your spreadsheet besides filling in the appropriate cells.<br><br>";
  
  var hmCombinedSheet = mainSpreadsheet.getSheetByName("Combined Results");
  var retMessage = "The script failed to share a copy of the spreadsheet with the following reviewers: <br>";
  var revFail = false;
  var newReviewerSpreadsheet;
  var formulaRange = hmCombinedSheet.getRange(3, numQuestions+2, 1000, 1).getA1Notation();  //range for linking sheets
  var hmAddMessage = "<br><br>There was an issue when trying to add the hiring manager as an editor in the following spreadsheets:<br>";
  var hmFail = false;

  var i;
  for(i = 0; i < reviewerEmails.length; i++){
    newReviewerSpreadsheet = mainSpreadsheet.copy(reviewerNames[i] + "'s Spark Hire Interview Review Sheet - " + positionTitle);
    newReviewerSpreadsheet.deleteSheet(newReviewerSpreadsheet.getSheetByName("Combined Results"));

    hmCombinedSheet.getRange(3, 2 + i).setFormula('=IMPORTRANGE("' + newReviewerSpreadsheet.getUrl() + '","' + formulaRange + '")');

    hmCombinedSheet.getRange(2, 2+i).setValue(reviewerNames[i]);
    try{
      newReviewerSpreadsheet.addEditor(reviewerEmails[i]);
      candidateList.addEditor(reviewerEmails[i]);
      MailApp.sendEmail({to:reviewerEmails[i], subject: reviewerEmailSubject, htmlBody: reviewerEmailBody + "<a href=\"" + candidateList.getUrl() + "\">Candidate List</a><br>" + "<a href=\"" + newReviewerSpreadsheet.getUrl() + "\">Your Spreadsheet</a>"});
    }
    catch(e){
      //most likely due to a reviewer email being typed wrong in the form
      revFail = true;
      retMessage += reviewerNames[i] + ", " + reviewerEmails[i] + ": <a href=\"" + newReviewerSpreadsheet.getUrl() + "\">" + reviewerNames[i] + "'s Spreadsheet </a><br>";
      //add the form user to the spreadsheet as an editor so that they can share it with the reviewer
      newReviewerSpreadsheet.addEditor(formUserEmail);
      candidateList.addEditor(formUserEmail);
    }
    
    try{
      newReviewerSpreadsheet.addEditor(hiringManagerEmail);
    }
    catch(e){
      //most likely due to hm email being typed wrong in the form
      newReviewerSpreadsheet.addEditor(formUserEmail);
      hmAddMessage += "<a href=\"" + newReviewerSpreadsheet.getUrl() + "\">" + reviewerNames[i] + "'s Spreadsheet </a><br>";
      hmFail = true;
    }

  }
  //link hm's scores to their combined results sheet
  hmCombinedSheet.getRange(3,2+i).setFormula('=ArrayFormula(Sheet1!' + formulaRange + ')');

  if(revFail == true){
    retMessage += "For each of these reviewers, please go into the spreadsheet that was created for them and the candidate list, manually share these with them, and email them a link to their spreadsheet and the candidate list. Once you have done this, please remove yourself from each reviewer's spreadsheet and the candidate list. Please note that this was most likely due to a typo when filling out the reviewer emails in the google form. The names and emails are displayed as they were filled out in the form. Any reviewer not displayed successfully received an email and their spreadsheet."
  }
  else{
    retMessage = "Spreadsheets were successfully created and shared with each reviewer. Each reviewer was sent an email informing them that they have been included as a reviewer. This email also contains links to their spreadsheet and the candidate list."
  }

  if(hmFail == true){
    retMessage += hmAddMessage + "For each of these spreadsheets, please manually add the hiring manager as an editor and then remove yourself.";
  }

  return retMessage;

}




function addEmailHm(hiringManagerEmail, formUserEmail, mainSpreadsheet, candidateListSpreadsheet, positionTitle){
  var fail = false;
  try{
      mainSpreadsheet.addEditor(hiringManagerEmail);
      candidateListSpreadsheet.addEditor(hiringManagerEmail);
      MailApp.sendEmail({
        to: hiringManagerEmail, 
        subject: "You have been added as the hiring manager for the \"" + positionTitle + "\" position.", 
        htmlBody: "Hello,<br><br> You are receiving this email to inform you that you have been included as the hiring manager in the hiring process for the \"" + positionTitle + "\" position. You should have been added as an editor of a spreadsheet where you can score applicants' responses to interview questions. This spreadsheet will be automatically populated with candidate first and last names as they are added to the canidate list spreadsheet for the position. On your spreadsheet, there is also a \"Combined Results\" sheet that is linked to spreadsheets belonging to the reviewers for this position. As they score applicant responses, your \"Combined Results\" sheet will automatically be populated with their scores. Attached below are the links to your spreadsheet and the candidate list spreadsheet. Do not make any modifications to your spreadsheet besides filling in the appropriate cells.<br><br> " + "<a href=\"" + candidateListSpreadsheet.getUrl() + "\">Candidate List</a><br>" + "<a href=\"" + mainSpreadsheet.getUrl() + "\">Your Spreadsheet</a>"
      });
    }
    catch(e){
      //most likely due to the email being typed wrong in the form
      fail = true;
      mainSpreadsheet.addEditor(formUserEmail);
      candidateListSpreadsheet.addEditor(formUserEmail);
    }

    if(fail == true){
      return "<br><br>There was an issue when trying to email and share the spreadsheets with the hiring manager. Please manually share their spreadsheet and the candidate list with them, and send them an email with links to their spreadsheet and the candidate list. Once you have done this, please remove yourself from their spreadsheet and the candidate list. Attached below is the link to their spreadsheet. <br><a href=\"" + mainSpreadsheet.getUrl() + "\">Hiring Manager's Spreadsheet</a>";
    }
    else{
      return "<br><br>A spreadsheet was successfully created and shared with the hiring manager. The hiring manager was sent an email informing them that they have been included as the hiring manager. This email also contains links to their spreadsheet and the candidate list.";
    }
}



function setUpTrigger() {
  ScriptApp.newTrigger('main')
  .forForm(' Google Form ID here')
  .onFormSubmit()
  .create();
}


