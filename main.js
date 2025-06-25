/**
 * Creates or updates an event in the user's default calendar from data inserted in a spreadsheet.
 * @see https://developers.google.com/calendar/api/v3/reference/events/insert
 * @see https://developers.google.com/calendar/api/v3/reference/events/update
 */

const calendarId = 'ADD YOU CALENDAR ID HERE'; // Replace with your calendar ID
const sheet = SpreadsheetApp.getActiveSheet(); // Get the active sheet in the spreadsheet
const events = sheet.getRange('A2:H1000').getValues(); // Get the values from the range A2:H1000 in the active sheet, which contains the event data

function createorUpdateEvents() {
  let event;

  for (i = 0; i < events.length; i++) {
    const shift = events[i];
    const eventID = shift[0];
    const eventsubject = shift[1];
    const startTime = shift[2];
    const endTime = shift[3];
    const description = shift[4];
    const color = shift[5];

    if (eventID !== undefined && eventsubject !== undefined && description !== undefined && color !== undefined && startTime instanceof Date && endTime instanceof Date) {
      const event = {
        id: eventID,
        summary: eventsubject,
        description: description,
        start: {
          dateTime: startTime.toISOString(),
          timeZone: 'America/Sao_Paulo',
        },
        end: {
          dateTime: endTime.toISOString(),
          timeZone: 'America/Sao_Paulo',
        },
        colorId: color,
      };

      try {
        let createOrUpdate;
        if (event.id) {
          createOrUpdate = Calendar.Events.update(event, calendarId, eventID);
        } else {
          createOrUpdate = Calendar.Events.insert(event, calendarId);
        }
      } catch (e) {
        if (e.message && e.message.indexOf('Not Found') !== -1) {
          createOrUpdate = Calendar.Events.insert(event, calendarId);
        } else {
          console.error('Error:', e);
        }
      }
    }
  }
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
    .addItem('Update Calendar', 'createorUpdateEvents') // Add an item to the menu that triggers the function to create or update events in the calendar
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
