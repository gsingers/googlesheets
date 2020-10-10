/***
 * Apache Software License 2.0
 *
 * In September of 2020, I combined work I had been doing with Igor Shindel as an executive coach and Cal Newport's Deep Work" book into a spreadsheet and these scripts to make
 * it easier to collect data about meetings as well as my day (since I book almost all of my time, per Deep Work), how effective they were, what my energy level was, the depth of my work
 * etc. Plus, I'm a cheapskate and don't want to pay for one of those fancy apps, but I also didn't want to hand copy over my calendar to a journal.
 *
 * To use this script, create a Google Sheet with the following columns:
 * Date;	Start time;	End time; 	Meeting, call, solo work description;	Who was involved;	"How was your energy? 1=very low to 5=very high";
 * "Postive or negative energy? -5 very negative to +5 very positive";	Depth of work? 1 for shallow, 5 for deep;	Notes (e.g. why?)
 *
 * Setup a calendar item to remind you to harvest the values every day at the end of the working day.  After a few weeks, try to identify where the patterns lay.
 *
 * Install this script usin the Tools-Script Editor.  The script will fill in all the columns up to, but not including, the "How was your Energy" column.  You gotta do that yourself!
 *
 * Assumes Google Calendar is in use.  Does not work for other calendars.
 *
 *
 * Props to Igor Shindel of Results Architect for the journaling approach and Spreadsheet template and Deep Work by Cal Newport for the idea of systematically tracking the deep work one is doing.
 *
 *
 * TODO:
 * - Probably some other smart things we could do to slice and dice the data or account for unscheduled time.
 *
 * - Might make sense to prepend at the top instead of the bottom so you don't have to scroll.
 *
 *
 */

function populateToday() {
  var today = new Date();
  populateDate(today);
}

function populateDate(theDate){
  var events = CalendarApp.getDefaultCalendar().getEventsForDay(theDate);
  Logger.log('Number of events: ' + events.length);
  var sss = SpreadsheetApp;
  var activeSpreadsheet = sss.getActive();
  var ss  = activeSpreadsheet.getActiveSheet();
  //var lastContentRow = ss.getLastRow();
  for (i =0; i < events.length; i++){
    if (events[i].isAllDayEvent() == false && events[i].getTitle().startsWith("End of Day") == false){
      //Logger.log(events[i].getTitle());
      var guests = events[i].getGuestList(true);//add in the people on the invite
      var whoInvolved = "";
      for (j =0; j < guests.length; j++){

        if (guests[j].getName() != null && guests[j].getName() != ""){
          if (guests[j].getName().startsWith("HQ-") == false){
            whoInvolved += guests[j].getName() + "\n";
          }
        } else{
          whoInvolved += guests[j].getEmail() + "\n";
        }
      }
      var formattedDate = formatDate(theDate);
      Logger.log(formattedDate);
      var rowContents = [formattedDate, formatTime(events[i].getStartTime()), formatTime(events[i].getEndTime()), events[i].getTitle(), whoInvolved];
      Logger.log(rowContents);
      ss.appendRow(rowContents);
    }
//    ss.appendRow([formatDate(today), ]);
  }
}

function populateForDate(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('What date?');
  if (response.getSelectedButton() == ui.Button.OK) {
    Logger.log('The date is %s.', response.getResponseText());
    var theDate = new Date(response.getResponseText());
    populateDate(theDate);
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function formatDate(date){
  return (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear()
}

function formatTime(date){
  return date.toTimeString();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Scripts")
   .addItem("Populate Today's Meetings", 'populateToday')
   .addItem("Populate for a date", 'populateForDate')
      .addToUi();
  ScriptApp.newTrigger('populateToday')
      .timeBased().everyDays(1)
      .atHour(17)
      .create();
}


