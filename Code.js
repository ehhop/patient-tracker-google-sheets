/////////////////////
//
// Patient Tracker Sheet v0.1
// By Ryan Neff
// EHHOP Tech
//
// Last updated: 10/28/17
//
// Usage: This runs in the background in a Google Sheets document 
// to track changes to the patient tracker call sheet automatically.
//
// If not working: you may also need to set up the trigger to fire in Google Sheets by clicking
// "Edit" > "Current project's triggers" > onEdit as in this image: https://imgur.com/x1H7uMf
//
// Public template (view only, no PHI):
// https://docs.google.com/spreadsheets/d/1wgs4391wcmzrWvFSggBL81YJSQmNaGzGfn3CxfeLH_0/edit?usp=sharing
//
/////////////////////

count = 0;

var s = SpreadsheetApp.getActiveSpreadsheet(); //get the current sheet
var timezone = Session.getScriptTimeZone(); //get the timezone (required for proper time formatting)
var sheet = s.getSheetByName("Callsheet"); //this is the main patient tracker sheet
var timesheet = s.getSheetByName("Timers - DO NOT EDIT"); //this is a second spreadsheet that has three columns: start time, stop time, and in appointment time
var offset = 16; // this is the # of columns on the callsheet before the "Main" column
var rowoffset = 1; // this is the number of rows to skip in the timesheet
var columns = 9;// this is the number of timers columns
var multiple = 3; // number of timers columns per column (start time, stop time, and in appointment time)
var checkincolumn = 11; //this is the column that says whether or not the patient has checked in to clinic

function onEdit(){ //this gets triggered every time the sheet gets edited
  var r = sheet.getActiveCell(); //get the cell that was edited in the document
  if((r.getColumn() > offset) && (r.getColumn() <= offset+columns)){ // only look at cells that we care about tracking the time of
    var column = r.getColumn(); //get the column of the edited cell
    var row = r.getRow()+rowoffset; //get the row of the edited cell
    var setting = r.getValue(); //get what the edited cell was changed to
    var time = new Date(); //create an empty datetime object
    var time = Utilities.formatDate(time, timezone, "HH:mm"); // Format date and time as hours and minutes
    if((setting == "Being Seen") || (setting=="In Pt Room")||(setting=="In Svc Room")){ // these are start words in the patient tracker sheet
      timesheet.getRange(row,(column-offset-1)*multiple+1).setValue(time); // put in time in start cell
      timesheet.getRange(row,(column-offset-1)*multiple+2).clear(); // reset stop cell (in case someone restarted an appt)
    };
    if((setting=="Done")||(setting=="Attending")||(setting=="With Attending")){ // these are stop words
      if(timesheet.getRange(row,(column-offset-1)*multiple+2).getValue()==""){ // only stop the appointment if not already stopped
        timesheet.getRange(row,(column-offset-1)*multiple+2).setValue(time); // put in time in stop cell
      };
    };
    if(setting==""){ //clear timers if the edited cell in the patient tracker is cleared
      timesheet.getRange(row,(column-offset-1)*multiple+1).clear();
      timesheet.getRange(row,(column-offset-1)*multiple+2).clear();
    };
  };
  if (sheet.getRange(r.getRow(),checkincolumn).getValue()==""){ //clear all timers from that row if the check in column is cleared (useful when starting a new clinic week)
      for(i=0; i < columns; i++){
        timesheet.getRange(row,i*multiple+1).clear(); //clear timers
        timesheet.getRange(row,i*multiple+2).clear(); //clear timers
      }
  };
};

function listenReady() { //this is an event listener that watches for changes in the patient tracker sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Callsheet");
  var range = sheet.getRange(2, 11, 103); //these are the cells we are interested in
  var values = range.getValues();
  onEdit();
}

function clearTimers(){ //this is a custom function to manually clear all timer columns
  for(i=0; i < columns; i++){
    timesheet.getRange(3,i*multiple+1,100).clear();
    timesheet.getRange(3,i*multiple+2,100).clear();
  }
}

function onOpen() { //this adds a menu to the top of the screen (after File, Edit, View...) with our clear timers function
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('**EHHOP Tech**') //create the menu title
      .addItem('Clear Timers', 'clearTimers') // create the menu button for the clear timers function
      .addToUi();
}
