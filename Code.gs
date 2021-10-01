function sendEmail() {  
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  var row = findCurrentWeekToSend(ss)-4;

  var message = "Hey Christine, here are the counts for the UTM runs: \n";
  for (var i = 0; i<5; i++){
    const template = HtmlService.createTemplateFromFile('body');
    var runs = {
      date: Utilities.formatDate(ss.getRange(row +i, 2).getValue(), (new Date).getTimezoneOffset(), 'MMMM dd'),
      oneA: ss.getRange(row +i, 3).getValue(),
      oneB: ss.getRange(row +i, 4).getValue(),
      twoA: ss.getRange(row +i, 5).getValue(),
      twoB: ss.getRange(row +i, 6).getValue(),
      threeA: ss.getRange(row +i, 7).getValue(),
      threeB: ss.getRange(row +i, 8).getValue()
    }
    template.sheet = runs;
    message += template.evaluate().getContent();
  }
  MailApp.sendEmail({
    to: "xxx.xxx@xxx.com", 
    subject: "UTM Counts", 
    htmlBody: message
    });
}
function findCurrentWeekToSend(ss){
  var ss = ss;
  var date = new Date();
  var timeZone = date.getTimezoneOffset();
  var dayString = Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
  for (var row = 3; row <= ss.getLastRow(); row++){
    var dateCell = ss.getRange(row,2).getValue();
    if(dateCell != ""){
      var dayString2 = Utilities.formatDate(dateCell, timeZone,'yyyy-MM-dd');
     if(dayString2 == dayString){
       return row;
     }
    }
  }
}
