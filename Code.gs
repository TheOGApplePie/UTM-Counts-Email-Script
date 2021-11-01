function sendEmail() {  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = findToday(ss)-4;

   var message = "Hey ****, here are the counts for the UTM runs: \n";
   var timeRow = row;

   while(ss.getRange(timeRow, 2).getValue()!= "TIME"){
    
     timeRow--;
   }
   for (var i = 0; i<5; i++){
      const template = HtmlService.createTemplateFromFile('body');
      runs = {
        date: Utilities.formatDate(ss.getRange(row +i, 2).getValue(), (new Date).getTimezoneOffset(), 'MMMM dd'),
        timeOneA: Utilities.formatDate(ss.getRange(timeRow, 3).getValue(),"GMT-500", 'HH:mm'),
        timeOneB: Utilities.formatDate(ss.getRange(timeRow, 4).getValue(),"GMT-500", 'HH:mm'),
        timeTwoA: Utilities.formatDate(ss.getRange(timeRow, 5).getValue(),"GMT-500", 'HH:mm'),
        timeTwoB: Utilities.formatDate(ss.getRange(timeRow, 6).getValue(),"GMT-500", 'HH:mm'),
        timeThreeA: Utilities.formatDate(ss.getRange(timeRow, 7).getValue(),"GMT-500", 'HH:mm'),
        timeThreeB: Utilities.formatDate(ss.getRange(timeRow, 8).getValue(),"GMT-500", 'HH:mm'),
        oneA: ss.getRange(row +i, 3).getValue(),
        oneB: ss.getRange(row +i, 4).getValue(),
        twoA: ss.getRange(row +i, 5).getValue(),
        twoB: ss.getRange(row +i, 6).getValue(),
        threeA: ss.getRange(row +i, 7).getValue(),
        threeB: ss.getRange(row +i, 8).getValue()
      };
      template.sheet = runs;
      message += template.evaluate().getContent();
   }
   MailApp.sendEmail({
     to: "xxx@xxx.com", 
     subject: "UTM Counts", 
     htmlBody: message
     });

}
/**
   ss - The current active spreadsheet
   Go through the entire spreadsheet and find the row corresponding to today
   Then find the last Friday thats happened. If today was a friday, then today 
   would be the last friday to have happened
   */
 function findToday(ss){
   var ss = ss;
   var todayDate = new Date();
   var timeZone = todayDate.getTimezoneOffset();
   var todayDateString = Utilities.formatDate(todayDate, timeZone, 'yyyy-MM-dd');
   var startRow = Math.floor((todayDate - new Date(2021,08,06))/(1000*60*60*24));
   for (var row = startRow; row <= ss.getLastRow(); row++){
     var cellDate = ss.getRange(row,2).getValue();
     if(cellDate != ""&& cellDate!="TIME"){
       var cellDateString = Utilities.formatDate(cellDate, timeZone,'yyyy-MM-dd');
      if(cellDateString == todayDateString){
        while(cellDate.getDay() != 5){
          row--;
          cellDate = ss.getRange(row,2).getValue();
        }
        Logger.log(row);
        return row;
      }
     }
   }
 }
