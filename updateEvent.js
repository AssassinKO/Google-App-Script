function onOpen() {
    ScriptApp
      .newTrigger('updateFunction')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onChange();
}
function updateFunction(e) {
  var spreadsheet = SpreadsheetApp.getActiveSheet();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets()[0];
  var row = sheets.getLastRow();

  var range = spreadsheet.getActiveRange();
  var row1 = range.getLastRow();
  if(row1 <= row) {
    var sh = SpreadsheetApp.getActive().getActiveSheet();
    var data = sh.getDataRange().getValues();
    for(n = 0; n <= data.length; n ++){
      if(n == row1-1) {
        var time_1 = new Date(data[n][12]);
        var time_2 = new Date(data[n][13]);
        var date_1 = new Date(data[n][11]);
        var hrs = time_1.getHours();
        var min = time_1.getMinutes();
        var sec = time_1.getSeconds();

        var hrs1 = time_2.getHours();
        var min1 = time_2.getMinutes();
        var sec1 = time_2.getSeconds();

        var dateAndTime = new Date(date_1).setHours(hrs,min,sec,0);
        var dateAndTime_2 = new Date(date_1).setHours(hrs1,min1,sec1,0);
        }
    }

        dateAndTime -= 69540000;
        dateAndTime_2 -= 69540000;
      var calendar = CalendarApp.getCalendarById("Stewart");
      var events = calendar.getEvents(new Date(dateAndTime), new Date(dateAndTime_2));
      for(var i = 0; i< events.length; i++) {
      events[i].deleteEvent();
      }


    var rng = spreadsheet.getRange(row1, 1, 1, 50);
    var rangeArray = rng.getValues();
    for(x = 0; x < rangeArray.length; x++) {
      var sheet = rangeArray[x];
      var firstName = sheet[6];
      var lastName = sheet[7];
      var jobId = sheet[4];

      var guestNumber = sheet[5];
      var teamEmail = sheet[41];
      var jobMethod = sheet[19];
      var jobType = sheet[17];
      var frequency = sheet[10];
      var access = sheet[20];
      var parking = sheet[21];

      var phone = sheet[9];
      var email = sheet[8];
      var team = sheet[33];
      var teamPhone = sheet[40];
    }

    var eventCal = CalendarApp.getCalendarById("Stewart");
    var signup = spreadsheet.getRange("AH" + row1).getValue();
    if(row1 <= row) {
    if(signup == ""){
      eventCal.createEvent((firstName + '     ' + lastName + '     #' + jobId), new Date(dateAndTime), new Date(dateAndTime_2), {description: jobMethod + ' Job\r' + jobType + '\\r' + frequency + '\r' + access + '\r' + parking + '\r\r' + firstName + '  ' + lastName +    '\r' + phone + '\r' + email});
    }else{
        eventCal.createEvent((firstName + '     ' + lastName + '     #' + jobId), new Date(dateAndTime), new Date(dateAndTime_2), {description: guestNumber+ " guest" + '\r\r' + teamEmail + '\r\r' + jobMethod + ' Job\r' + jobType + '\r' + frequency + '\r' + access + '\r' + parking + '\r\r' + firstName + '  ' + lastName +    '\r' + phone + '\r' + email + '\r\r' + team + '\r' + teamPhone});
    }
    }
  }
  }