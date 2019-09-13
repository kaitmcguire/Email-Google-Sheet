function sendEmail() {
// This script will send a daily email of this spreadsheet to the email(s) I choose

// This script will automatically send the email to the email of the owner of this sheet, but you can specify other people (or multiple people) with var specialEmail

  var specialEmail = "test@test.com";



// This specifies the sheet you want sent to your email

  var sheetName = "Sheet1";
 
// This sets the subject

  var subject = "Subject of Email I Want This to Send";

// You can set the alignment method for the cells in the email, either all the same or custom and then set each column

  var alignment = "custom";

  var customAlign = ["center","left","left","left"];
  
// The email can display call rows and columns OR you can set the number of rows and columns manually
// To display all rows and column, use

  var rowDisplay = "all";
  var columnDisplay = 4;

// this is the actual script

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  var range = sheet.getDataRange();
  var values = range.getDisplayValues();
  
  // this gets an array of font weights, normal or bold, for each cell
  var weights = range.getFontWeights();
  
  // this is used to determine if a row is blank
  // it checks column B for a category name
  var blankRowCount = 0;
   
  if (rowDisplay == 'all')
      { rowDisplay = range.getLastRow(); }
      
  if (columnDisplay == 'all')
      { columnDisplay = range.getLastColumn(); }
  
  var message = '';

//    
  // this starts the table for the cells
  message = message + "<table rules='all' style='border-color: #666;' cellpadding='5'>\n";
  
  for (var i = 0; i < rowDisplay; i++ ) 
  {  
    {
      message = message + "<tr>\n";
    
      for (var j = 0; j < columnDisplay; j++ ) 
      {
      
        if (customAlign[j] != 'skip')
        {
          message = message + "<td style='text-align: ";
        
          if (alignment == 'custom')
          {
             message = message + customAlign[j];            
          }
          else
          {
             message = message + alignment;
          }
             
          message = message + "; font-weight: " + weights[i][j] + ";'>" + values[i][j]   + "</td>\n";
        }  // ends if no skip
       } // ends the for loop for each column
      
      message = message + "</tr>\n"; 

    }
  }
  message = message + "</table>\n";
  // use the users email or the other Email
    if (specialEmail == 'none')
     {var email = Session.getActiveUser().getEmail();}
    else
     {var email = specialEmail;}
  
  // send the email
  MailApp.sendEmail({
     to: email,
     subject: subject,
     htmlBody: message});    
}
