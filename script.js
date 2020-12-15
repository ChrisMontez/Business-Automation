function sendHeadshotNotice() {
  
  
  
  
    const getData =  async(name) => {  
      const data = {
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
           } 

        return data(name);                 
      }
  
   

  getData().then(() => {
          
         if (headshots === 'TRUE' && offerSent === 'FALSE') {
           var templ = HtmlService.createTemplateFromFile('head_shots');
           var msg = templ.evaluate().getContent();
           MailApp.sendEmail(email, 'Your Business Headshot is Ready!', msg,{ noReply: false, htmlBody: msg })
           s.getRange(actualRow, 5 /* offerSentCol */, 1).setValues([['TRUE']]);
           numSent++;
           s.getRange(actualRow,6).setValue(new Date());
     
         }
      
        SpreadsheetApp.flush();
          if (numSent > 0) {
          ss.toast('Sent ' + numSent + '  new letters.');
          } else {
          ss.toast('No letters sent');
          }          

   })

  
}  
  
  
  
  
  
  
  
  
  