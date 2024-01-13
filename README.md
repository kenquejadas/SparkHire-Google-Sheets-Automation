# SparkHire-Google-Sheets-Automation
Automation of HR process to create spreadsheets of candidates SparkHire interview data upon data entry into a Google Form

## Background
This project was done in collaboration with one other student automation specialist in the Student Automation Team of the University of Minnesota Office of Information Technology.
I have received written permission from their HR team to describe the process and showcase the script on my public portfolio.

## Process
1. After each candidate interview, a recruiter fills out the SparkHire Review Google Form. The data input includes:
    - Position Title
    - Hiring Manager Name and Email
    - Interview Questions
    - Reviewer Names and Emails
2. Upon form submission, 3 actions are done:
    - Form answers automatically entered onto a Google Sheet for Google Form submission records
    - Script creates a Candidate List spreadsheet
        - Hiring Manager and Reviewers manually enter candidate names onto the spreadsheet
        - This Candidate List is linked to the review spreadsheets that the script also creates
    - Script creates review spreadsheets for the Hiring Manager and each reviewer
          - First column is auto-populated with candidates from the Candidate List spreadsheet
          - The next columns are scoring for answers to interview questions and total score
3. Hiring Manager and each Reviewer receive an email with links to the generated review spreadsheets and candidate list
  
### Process Diagram
![Process Diagram](https://github.com/kenquejadas/SparkHire-Google-Sheets-Automation/blob/main/ProcessDiagram.png)

## Technical Approach
![Approach DFD](https://github.com/kenquejadas/SparkHire-Google-Sheets-Automation/blob/main/AutomationApproach.png)

## Script Documentation
### main()
 - Main function, gets triggered by a form submission. 
 - Reads the last row in the spreadsheet tied to the google form to get form response data and populate variables. 
 - Has a lock in case multiple people try to submit a form response at the same time. 
 - If the number of reviewer emails is not equal to the number of reviewer names, the script sends an email to the form user and stops running
 - Creates the candidate list spreadsheet
 - Calls createSpreadsheet(), createShareLinkReviewerSpreadsheets(), and addEmailHm()
 - Sends email to form user with results of running the script

### createSpreadsheet()
 - Creates the hiring manager's spreadsheet and returns it
 - Does all the cell formatting, highlighting, and adds the interview questions
 - Adds the sum formula to sheet1 and average formula to combined results sheet
 - Links both sheets to the candidate list
 - Sets up cell protections

### createShareLinkReviewerSpreadsheets()
 - Does bulk of the work
 - Returns a string containing part of the message to be included in the form user email 
 - Makes copies of the main spreadsheet and deletes the combined results sheet from each copy
 - Attempts to add the hiring manager and a reviewer to each spreadsheet, if it fails to do so it adds the form user
 - Links each reviewer spreadsheet to the hiring manager's combined results sheet to automatically add their scores
 - Hiring manager's scores also get linked to the combined results sheet in this function

### addEmailHm()
 - Tries to add the hiring manager to their spreadsheet and email them
 - Adds the form user if it fails to add the hiring manager
 - Returns a string containing part of the message to be included in the form user email

### setUpTrigger() 
 - Never gets called and should not be run
 - Used to set up the trigger that causes the script to run in response to a form submission
 - Had the HR account run this once when they first got the script to set up the trigger under the HR account so emails send from them and
   they are the owner of each created spreadsheet
