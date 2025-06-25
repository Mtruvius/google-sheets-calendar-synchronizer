/**
 * Creates or updates an event in the user's default calendar from data inserted in a spreadsheet.
 * @see https://developers.google.com/calendar/api/v3/reference/events/insert
 * @see https://developers.google.com/calendar/api/v3/reference/events/update
 */

const calendarId = 'ADD YOUR CALENDAR ID HERE'; // Replace with your calendar ID
const sheet = SpreadsheetApp.getActiveSheet(); // Get the active sheet in the spreadsheet
const events = sheet.getRange('A2:H1000').getValues(); // Get the values from the range A2:H1000 in the active sheet, which contains the event data

/*
 * This function iterates through the events data, checks if an event ID exists, and either updates the existing event or creates a new one. It also handles deletion of events based on a flag in the data.
 * @see https://developers.google.com/apps-script/reference/calendar/calendar-app
 * @see https://developers.google.com/apps-script/reference/calendar/calendar
 * @see https://developers.google.com/apps-script/reference/calendar/event
 * @see https://developers.google.com/apps-script/reference/spreadsheet/sheet
 */
function createOrUpdateEvents() {
  let theEvent; // Declare the event variable outside the loop

  var calendarResponse = Calendar.Calendars.get(calendarId); // Get the calendar details using the calendar ID
  var defaultTimeZone = calendarResponse.timeZone; // Get the default time zone of the calendar

  const calendar = CalendarApp.getCalendarById(calendarId); // Get the calendar by its ID
  const startDate = new Date(events[0][2]); // Get the start date from the first event in the data

  const fCsvEvents = events.filter(value => value[1] != ''); // Filter out empty rows from the events data
  const endIndex = fCsvEvents.length - 1; // Get the index of the last event in the filtered data
  const endDate = new Date(events[endIndex][3]); // Get the end date from the last event in the data

  fCsvEvents.forEach(event => {
    // Iterate through each event in the filtered events data
    const eventID = event[0]; // Get the event ID from the first column of the event data
    const eventsubject = event[1]; // Get the event subject from the second column of the event data
    const startTime = event[2]; // Get the start time from the third column of the event data
    const endTime = event[3]; // Get the end time from the fourth column of the event data
    const description = event[4]; // Get the description from the fifth column of the event data
    const color = event[5]; // Get the color from the sixth column of the event data
    const guests = event[6]; // Get the guests from the seventh column of the event data
    var deleteOpt = event[7]; // Get the delete option from the eighth column of the event data

    if (deleteOpt) {
      // If the delete option is set, delete the event from the calendar
      calendar.getEventById(eventID).deleteEvent(); // If the delete option is set, delete the event from the calendar
      return; // Skip to the next iteration if the event is marked for deletion
    }

    if (eventID !== undefined && eventsubject !== undefined && description !== undefined && color !== undefined && startTime instanceof Date && endTime instanceof Date) {
      // Check if the event ID, subject, description, color, start time, and end time are defined and valid
      const theEvent = {
        // Create an event object with the necessary properties
        id: eventID, // Use the event ID from the data
        summary: eventsubject, // Use the event subject from the data
        description: description, // Use the description from the data
        start: {
          // Set the start time of the event
          dateTime: startTime.toISOString(), // Convert the start time to ISO string format
          timeZone: defaultTimeZone, // Use the default time zone of the calendar
        },
        end: {
          // Set the end time of the event
          dateTime: endTime.toISOString(), // Convert the end time to ISO string format
          timeZone: defaultTimeZone, // Use the default time zone of the calendar
        },
        colorId: color, // Use the color from the data
      };
      /**
       * If the event ID exists, update the existing event; otherwise, create a new event.
       **/
      try {
        // Try to create or update the event
        let createOrUpdate; // Declare the variable to hold the result of the create or update operation
        if (theEvent.id) {
          // If the event ID exists, update the existing event
          createOrUpdate = Calendar.Events.update(theEvent, calendarId, eventID); // Update the event using the Calendar API
        } else {
          // If the event ID does not exist, create a new event
          createOrUpdate = Calendar.Events.insert(theEvent, calendarId); // Insert the new event using the Calendar API
        }
      } catch (e) {
        // Catch any errors that occur during the create or update operation
        if (e.message && e.message.indexOf('Not Found') !== -1) {
          // If the error message indicates that the event was not found
          createOrUpdate = Calendar.Events.insert(theEvent, calendarId); // Attempt to insert the event again
        } else {
          // If the error is not related to the event not being found
          console.error('Error:', e); // Log the error to the console
        }
      }
    }
  });

  importCalendarEventsToSheet({start: startDate, end: endDate}); // Call the function to import calendar events to the sheet with the specified start and end dates
}

/*
 * Display the dialog to select the date range
 * @see https://developers.google.com/apps-script/guides/dialogs
 * @see https://developers.google.com/apps-script/guides/html/communication
 */
function displayDatesDialog() {
  var ui = SpreadsheetApp.getUi(); // Get the user interface of the spreadsheet
  var html = HtmlService.createHtmlOutputFromFile('dateSelect').setWidth(350).setHeight(80); // Create an HTML output from the 'dateSelect' file, setting its width and height
  ui.showModalDialog(html, 'Date Range Selection'); // Show the dialog with the title 'Date Range Selection'
}

/*
 * Import calendar events to the sheet
 ** @param {Object} e - The event object containing the start and end dates selected in the dialog.
 */
async function importCalendarEventsToSheet(e) {
  const startDate = new Date(e.start); // declare startime obtained from the selected range in the dates dialog.
  const endDate = new Date(e.end); // declare endtime obtained from the selected range in the dates dialog.
  const calendar = CalendarApp.getCalendarById(calendarId); // Get the calendar by its ID
  const events = calendar.getEvents(startDate, endDate); // Get the events within the specified date range
  const data = []; // Initialize an array to hold the event data
  if (events.length > 0) {
    events.forEach(event => {
      const eventID = event.getId().split('@')[0]; // Extract the event ID before the '@' symbol
      const eventTitle = event.getTitle(); // Get the event title
      const startTime = event.getStartTime(); // Get the start time of the event
      const endTime = event.getEndTime(); // Get the end time of the event
      const description = event.getDescription(); // Get the description of the event
      const guests = undefined; // Get the guests of the event
      const color = event.getColor(); // Get the color of the event
      const deleteOpt = 'false'; // Placeholder for delete option
      data.push([eventID, eventTitle, startTime, endTime, description, color, guests, deleteOpt]); // Push the event data into the array
    });
    const numRows = data.length; // Get the number of rows in the data array
    const numCols = data[0].length; // Get the number of columns in the data array
    sheet.getRange(2, 1, numRows, numCols).setValues(data); // Set the values in the sheet starting from row 2, column 1
  } else {
    console.log('No events exist for the specified range'); // Log a message if no events are found
  }
}

/*
 * Creates a menu to access the functions in the Google Sheets UI
 * @see https://developers.google.com/apps-script/guides/menus
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi(); // Get the user interface of the spreadsheet
  ui.createMenu('Sync Data with Calendar') // Create a new menu in the spreadsheet UI
    .addItem('Import Events', 'displayDatesDialog') // Add an item to the menu that opens a dialog for date selection
    .addItem('Update Calendar', 'createOrUpdateEvents') // Add an item to the menu that triggers the function to create or update events in the calendar
    .addSeparator() // Add a separator in the menu
    .addSubMenu(ui.createMenu('About').addItem('Documentation', 'showDocumentation')) // Add a submenu with an item that shows documentation
    .addToUi(); // Add the menu to the user interface
}

/**
 * Displays a modal dialog with documentation link.
 * @see https://developers.google.com/apps-script/guides/html/communication
 */
function showDocumentation() {
  var htmlOutput = HtmlService.createHtmlOutput('<p>For more info, visit <a href="https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer" target="_blank">this link</a>.'); // Create an HTML output with a link to the documentation
  var ui = SpreadsheetApp.getUi(); // Get the user interface of the spreadsheet
  ui.showModalDialog(htmlOutput, 'Documentation'); // Show the dialog with the documentation link
}
