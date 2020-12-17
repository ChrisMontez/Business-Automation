

function addMenuItems() {

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send Letters')
  .addItem('Send Orientation Letters', 'sendNewLetters')
  .addItem('Send Next Steps', 'sendNextSteps')
  .addItem('Send Headshot Notice', 'sendHeadshotNotice')
  .addItem('Send Marketing Items', 'sendMarketingItems')
  .addToUi();
}


function onEdit(e) {
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  
   if(col === 1 && row > 1 && e.source.getActiveSheet().getName() === "Headshots"  ) {   
    e.source.getActiveSheet().getRange(row,4).setDataValidation(checkbox).setValue(false);
    e.source.getActiveSheet().getRange(row,5).setDataValidation(checkbox).setValue(false);
   }
   if(col === 1 && row > 1 && e.source.getActiveSheet().getName() === "Marketing Items"  ) {   
    e.source.getActiveSheet().getRange(row,4).setDataValidation(checkbox).setValue(false);
    e.source.getActiveSheet().getRange(row,5).setDataValidation(checkbox).setValue(false);
    e.source.getActiveSheet().getRange(row,6).setDataValidation(checkbox).setValue(false);
    e.source.getActiveSheet().getRange(row,7).setDataValidation(checkbox).setValue(false);
   }
  if(col === 1 && row > 1 && e.source.getActiveSheet().getName() === "Next Steps"  ) {   
    e.source.getActiveSheet().getRange(row,4).setDataValidation(checkbox).setValue(false);
   }
   if(col === 1 && row > 1 && e.source.getActiveSheet().getName() === "Orientation"  ) {   
    e.source.getActiveSheet().getRange(row,4).setDataValidation(checkbox).setValue(false);
   }
}





function sendNewLetters() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Sending Letters...');
  var s = ss.getSheetByName('Orientation');
  var lastRow = s.getLastRow();
  var range = s.getRange(2 /* startRow */, 1 /* startCol */, lastRow -1 /* numRows */, 8 /* numCols */ );
  
  var data = range.getDisplayValues();
  
  
  var numSent = 0;
  for (var i = 0; i < data.length; i++) {
    var actualRow = i + 2;
    var row = data[i];
    var lastname = row [0];
    var firstname = row[1];
    var email = row[2];
    var offerSent = row[3]
    if (offerSent != 'TRUE') {
      var templ = HtmlService
      .createTemplateFromFile('orientation_email');
      
      var msg = templ.evaluate().getContent();
      
      MailApp.sendEmail(email, 'Welcome to Integrity Real Estate Group', msg,{ noReply: true, htmlBody: msg })
      s.getRange(actualRow, 4 /* offerSentCol */, 1).setValues([['TRUE']]);
      s.getRange(actualRow,5).setValue(new Date());
      numSent++;
    }
  }
  
  SpreadsheetApp.flush();
  if (numSent > 0) {
    ss.toast('Sent ' + numSent + '  new letters.');
  } else {
    ss.toast('No letters sent');
  }
}


function sendNextSteps() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Sending Next Steps...');
  var s = ss.getSheetByName('Next Steps');
  var lastRow = s.getLastRow();
  var range = s.getRange(2 /* startRow */, 1 /* startCol */, lastRow -1 /* numRows */, 8 /* numCols */ );
  var data = range.getDisplayValues();
  var numSent = 0;
  for (var i = 0; i < data.length; i++) {
    var actualRow = i + 2;
    var row = data[i];
    var last_name = row [0];
    var first_name = row [1];
    var email = row[2];
    var offerSent = row[3]
    if (offerSent != 'Yes') {
      var templ = HtmlService
      .createTemplateFromFile('next_steps_email');
      var msg = templ.evaluate().getContent();
      MailApp.sendEmail(email, 'Your Next Steps - You\'re Well On Your Way!', msg,{ noReply: true, htmlBody: msg })
      s.getRange(actualRow, 4 /* offerSentCol */, 1).setValues([['TRUE']]);
      numSent++;
      s.getRange(actualRow,5).setValue(new Date());
    }
  }
  SpreadsheetApp.flush();
  if (numSent > 0) {
    ss.toast('Sent ' + numSent + '  new letters.');
  } else {
    ss.toast('No letters sent');
  }
}



  
function sendHeadshotNotice() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Sending Headshot completion notice...');
  var s = ss.getSheetByName('Headshots');
  var lastRow = s.getLastRow();
  var range = s.getRange(2 /* startRow */, 1 /* startCol */, lastRow -1 /* numRows */, 8 /* numCols */ );
  var data = range.getDisplayValues();
  var numSent = 0;
  for (var i = 0; i < data.length; i++) {
    var actualRow = i + 2;
    var row = data[i];
    var last_name = row [0];
    var first_name = row [1];
    var email = row[2];
    var headshots = row[3];
    var offerSent = row[4];
    var timeStamp = row[5]
    if (headshots === 'TRUE' && offerSent === 'FALSE') {
      var templ = HtmlService.createTemplateFromFile('head_shots');
      var msg = templ.evaluate().getContent();  
      MailApp.sendEmail(email, 'Your Business Headshot is Ready!', msg,{ noReply: false, htmlBody: msg })
      s.getRange(actualRow, 5 /* offerSentCol */, 1).setValues([['TRUE']]);
      numSent++;
      s.getRange(actualRow,6).setValue(new Date());
    }
  }

  
  SpreadsheetApp.flush();
  if (numSent > 0) {
    ss.toast('Sent ' + numSent + '  new letters.');
  } else {
    ss.toast('No letters sent');
  }  
  
} 


  
  


function sendMarketingItems() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Sending Marketing Items...');
  var s = ss.getSheetByName('Marketing Items');
  var lastRow = s.getLastRow();
  var range = s.getRange(2 /* startRow */, 1 /* startCol */, lastRow -1 /* numRows */, 8 /* numCols */ );
  var data = range.getDisplayValues();
  var numSent = 0;
  for (var i = 0; i < data.length; i++) {
    var actualRow = i + 2;
    var row = data[i];
    var lastname = row [0];
    var firstname = row [1];
    var email = row[2];
    var businesscard = row[3];
    var yardsign = row[4];
    var nametag = row[5];
    var offersent = row[6];
    var timestamp = row[7]
    if (businesscard === 'TRUE' && yardsign === 'TRUE' && nametag === 'TRUE' && offersent === 'FALSE') {
      
       
      var templ = HtmlService
      .createTemplateFromFile('marketing_items');

      var msg = templ.evaluate().getContent();
      
      MailApp.sendEmail(email, 'Your Next Steps - You\'re Well On Your Way!', msg,{ noReply: true, htmlBody: msg })
      s.getRange(actualRow, 7 /* offerSentCol */, 1).setValues([['TRUE']]);
      numSent++;
      s.getRange(actualRow,8).setValue(new Date());
     
    }
        
  }
  

  SpreadsheetApp.flush();
  if (numSent > 0) {
    ss.toast('Sent ' + numSent + '  Letter Containing Marketing-Items.');
  } else {
    ss.toast('No letters sent');
  }
  
}  


  

    




