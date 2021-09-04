function onOpen() {
  ScriptApp
    .newTrigger('firstFunction')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange();
}
function firstFunction(e) {
  var spreadsheet = SpreadsheetApp.getActiveSheet();

  var sh = SpreadsheetApp.getActive().getActiveSheet();
  var data = sh.getDataRange().getValues();
  data.shift();
  for(var n in data){
    var date_1 = new Date(data[n][11]);
    var time_1 = new Date(data[n][12]);
    var time_2 = new Date(data[n][13]);
    var hrs = time_1.getHours();
    var min = time_1.getMinutes();
    var sec = time_1.getSeconds();

    var hrs1 = time_2.getHours();
    var min1 = time_2.getMinutes();
    var sec1 = time_2.getSeconds();

    var dateAndTime = new Date(date_1).setHours(hrs,min,sec,0);
    var dateAndTime_2 = new Date(date_1).setHours(hrs1,min1,sec1,0);
  }

  dateAndTime -= 1380000;
  dateAndTime_2 -= 1380000;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var row = sheet.getLastRow();

  var range = spreadsheet.getActiveRange();
  var row1 = range.getLastRow();

    var rng = spreadsheet.getRange(row1, 1, 1, 50);
    var rangeArray = rng.getValues();
    for(x = 0; x < rangeArray.length; x++) {
      var sheet = rangeArray[x];
      var payStatus = sheet[0];
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
    console.log("signup", signup);
    if(payStatus != "" && row <= row1) {
    if(signup == ""){
      eventCal.createEvent((firstName + '     ' + lastName + '     #' + jobId), new Date(dateAndTime), new Date(dateAndTime_2), {description: jobMethod + ' Job\r' + jobType + '\r' + frequency + '\r' + access + '\r' + parking + '\r\r' + firstName + '  ' + lastName +    '\r' + phone + '\r' + email});
    }else{
        eventCal.createEvent((firstName + '     ' + lastName + '     #' + jobId), new Date(dateAndTime), new Date(dateAndTime_2), {description: guestNumber+ " guest" + '\r\r' + teamEmail + '\r\r' + jobMethod + ' Job\r' + jobType + '\r' + frequency + '\r' + access + '\r' + parking + '\r\r' + firstName + '  ' + lastName +    '\r' + phone + '\r' + email + '\r\r' + team + '\r' + teamPhone});
    }
    }
}