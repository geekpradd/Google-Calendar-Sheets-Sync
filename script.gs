function sync() {
    var spreadsheet = SpreadsheetApp.getActiveSheet();
    var calendarID = "pbora2000@gmail.com";
    var eventCal = CalendarApp.getCalendarById(calendarID);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("May");
  var num_columns =sheet.getLastColumn(); //column Index 

  Logger.log(num_columns)
  var firstColumn = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  
  for (i=2; i<num_columns; i+=2){
    var column_data = sheet.getRange(1, i, sheet.getLastRow()).getValues();
    var prev = "null";
    var start_times = [];
    var end_times = [];
    var titles = [];

    for (j=1; j<column_data.length; ++j){
      var str = column_data[j].toString();
     
      if (str === prev){
        if (j===(column_data.length - 1)){
          var date_current = new Date(column_data[0]);
          var time_current = new Date(firstColumn[j]);
          date_current.setHours(date_current.getHours() + time_current.getHours());
          end_times.push(date_current);
        }
        continue;
      }
        
      var date_current = new Date(column_data[0]);
      var time_current = new Date(firstColumn[j]);
      date_current.setHours(date_current.getHours() + time_current.getHours());
      date_current.setMinutes(time_current.getMinutes());
      if (j!==1){
        end_times.push(date_current);
      }
      start_times.push(date_current);
      titles.push(str);
      
      prev = str;

    }

    var entries = titles.length;

    for(k=0;k<entries;++k){
      Logger.log(start_times[k]);
      Logger.log(end_times[k]);
      
      Logger.log(titles[k]);
      if (end_times[k] < start_times[k]){
        end_times[k].setDate(end_times[k].getDate()+1);
      }
      var events = eventCal.getEvents(start_times[k], end_times[k]);
      var same = 0;
      for (l=0;l<events.length; ++l){
        if (events[l].getTitle()===titles[k]){
          same = 1;
        }
      }
      if (same === 0){
        eventCal.createEvent(titles[k], start_times[k], end_times[k]);
      }
      else {
        Logger.log("Already found");
      }
      
    }

    
  }
  
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar').addItem('Sync', 'sync').addToUi();
}