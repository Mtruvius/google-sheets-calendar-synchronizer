/****************************************
                * INIT *
*****************************************/
let settings; // Global variable to store settings
/**
 * Retrieves the settings for the Google Sheets Calendar Synchronizer.
 * @see https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#getactivespreadsheet
 * @see https://developers.google.com/apps-script/reference/spreadsheet/sheet#getactivesheet
 * @see https://developers.google.com/apps-script/reference/spreadsheet/range#getvalues
 * @see https://developers.google.com/apps-script/reference/properties/properties-service#getuserproperties
 * @see https://developers.google.com/apps-script/reference/calendar/calendar-app#getcalendarbyid(id)
 */
function GetSettings() {
    if (settings) return settings; // If settings are already defined, return them
    settings = { // Initialize settings object with default values
        dialogWidth: 350, // Width of the dialog in pixels
        dialogHeight: 90, // Height of the dialog in pixels
        spreadSheet: SpreadsheetApp.getActive(), // Get the active spreadsheet
        sheet: SpreadsheetApp.getActiveSheet(), // Get the active sheet in the spreadsheet
        events: SpreadsheetApp.getActiveSheet().getRange('A2:L1000').getValues(), // Get the values from the range A2:L1000 in the active sheet, which contains the event data
        calendarId: PropertiesService.getUserProperties().getProperty('CALID'), // Get the calendar ID from user properties
        defaultTimeZone: CalendarApp.getCalendarById(PropertiesService.getUserProperties().getProperty('CALID')).getTimeZone(), // Get the default time zone of the calendar
    };
    return settings; // Return the settings object
}

/**
 * This function is triggered when the Google Sheets document is opened. It creates a custom menu in the spreadsheet UI with options to import events, update the calendar, clear the sheet, and access documentation.
 *  @see https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#getui
 *  @see https://developers.google.com/apps-script/reference/spreadsheet/ui#createMenu(name)
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi(); // Get the user interface of the spreadsheet
    ui.createMenu('Manage Calendar') // Create a new menu in the spreadsheet UI
        .addItem('Import Events', 'FetchCalendarEvents') // Add an item to the menu that Fetchs the Calendar Events
        .addItem('Update Calendar', 'CreateOrUpdateEvents') // Add an item to the menu that triggers the function to create or update events in the calendar    
        .addSeparator() // Add a separator in the menu
        .addSubMenu(ui.createMenu('Settings').addItem('Add Calendar ID', 'ShowAddCalendarId')) // Add a submenu with an item that shows Add Calendar
        .addItem('Clear Sheet', 'ClearSheet') // Add an item to the menu that clears the sheet
        .addSeparator() // Add a separator in the menu
        .addSubMenu(ui.createMenu('Documentation').addItem('About', 'ShowAbout') // Add a submenu with an item that shows information about the script
            .addItem('Help', 'ShowDocumentation')) // Add a submenu with an item that shows documentation
        .addToUi(); // Add the menu to the user interface
}

/****************************************
            * IMPORT EVENTS *
*****************************************/
/**
 * This function is triggered when the user selects the "Import Events" option from the custom menu. It checks if a calendar ID exists, and if so, it displays a modal dialog for the user to select a date range for importing events from the calendar.
 */
function FetchCalendarEvents() {
    if (!CalendarIdExists()) return; // Check if a calendar ID exists, if not, display an error message and return
    const ui = SpreadsheetApp.getUi(); // Get the user interface of the spreadsheet
    var html = HtmlService.createTemplateFromFile('dateSelect').evaluate().setWidth(settings.dialogWidth).setHeight (settings.dialogHeight); // Create an HTML template from the file 'dateSelect' and set its width and height
    ui.showModalDialog(html, 'Date Range Selection'); // Show the dialog with the title 'Date Range Selection'  
}

/*
 * Add events to the sheet
 ** @param {Object} e - The event object containing the start and end dates selected in the dialog.
 */
async function AddEventsToSheet(e) {
    const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer  
    const sheet = settings.sheet; // Get the active sheet in the spreadsheet    
    try {   
        const startDate = new Date(e.start); // declare startime obtained from the selected range in the dates dialog.
        const endDate = new Date(e.end); // declare endtime obtained from the selected range in the dates dialog.
        const calendar = CalendarApp.getCalendarById(settings.calendarId); // Get the calendar by its ID
        const events = calendar.getEvents(startDate, endDate); // Get the events within the specified date range
        const data = []; // Initialize an array to hold the event data
        if (events.length > 0) { // Check if there are any events in the specified date range
            events.forEach(event => { // Iterate through each event in the events array
                const eventID = event.getId().split('@')[0]; // Extract the event ID before the '@' symbol
                const eventTitle = event.getTitle(); // Get the event title
                const startTime = event.getStartTime(); // Get the start time of the event
                const endTime = event.getEndTime(); // Get the end time of the event
                const isAllDay = event.isAllDayEvent(); // Get the end time of the event
                const description = event.getDescription(); // Get the description of the event
                const color = event.getColor(); // Get the color of the event
                const guests = event.getGuestList().map(g => `${g.getEmail()} (${g.getGuestStatus().toString().toLowerCase()})`).join(', '); // Get the guests of the event
                const myStatus = event.getMyStatus();
                const location = event.getLocation();
                const deleteOpt = 'false'; // Placeholder for delete option
                data.push([eventID, eventTitle, startTime, endTime, isAllDay, description, color, guests, myStatus, location, deleteOpt]); // Push the event data into the array
            });
            const numRows = data.length; // Get the number of rows in the data array
            const numCols = data[0].length; // Get the number of columns in the data array
            ClearSheet(); // Clear the existing content in the sheet
            sheet.getRange(2, 1, numRows, numCols).setValues(data); // Set the values in the sheet starting from row 2, column 1
            var startTimeColumn = sheet.getRange("C2:C"); // Get the range for the start time column
            startTimeColumn.setNumberFormat("dd/mm/YYYY HH:mm:ss"); // Set the number format for the start time column
            var endTimeColumn = sheet.getRange("D2:D"); // Get the range for the end time column
            endTimeColumn.setNumberFormat("dd/mm/YYYY HH:mm:ss"); // Set the number format for the end time column
        } else { // If there are no events in the specified date range
            ClearSheet(); // Clear the existing content in the sheet
        }
    }
    catch (e) { // Catch any errors that occur during the process
        DisplayError(e); // Display an error message if an error occurs
        return;
    }
}

/****************************************
            * CREATE & UPDATE EVENTS *
*****************************************/
/*
 * This function iterates through the events data, checks if an event ID exists, and either updates the existing event or creates a new one. It also handles deletion of events based on a flag in the data.
 * @see https://developers.google.com/apps-script/reference/calendar/calendar-app
 * @see https://developers.google.com/apps-script/reference/calendar/calendar
 * @see https://developers.google.com/apps-script/reference/calendar/event
 * @see https://developers.google.com/apps-script/reference/spreadsheet/sheet
 */
function CreateOrUpdateEvents() {
    if (!CalendarIdExists()) return; // Check if a calendar ID exists, if not, display an error message and return
    const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer
    const filteredCsvEv = settings.events.filter(value => value[1] != ''); // Filter out empty rows from the events data
    const startEndDates = GetStartEndDates(filteredCsvEv); // Get the start and end dates from the filtered events data
    const startDate = startEndDates[0]; // Get the start date from the first event in the data
    const endDate = startEndDates[1]; // Get the end date from the last event in the data

    filteredCsvEv.forEach(async (event) => {
        // Iterate through each event in the filtered events data
        const eventID = event[0]; // Get the event ID from the first column of the event data
        const eventTitle = event[1]; // Get the event subject from the second column of the event data
        const startTime = event[2]; // Get the start time from the third column of the event data
        const endTime = event[3]; // Get the end time from the fourth column of the event data
        const isAllDayEv = event[4]; // Get the all-day event flag from the fifth column of the event data
        const description = event[5]; // Get the description from the fifth column of the event data
        const color = event[6] > 0 ? event[6] : 8; // Get the color from the sixth column of the event data
        const guests = event[7].split(','); // Get the guests from the seventh column of the event data
        const myStatusStr = event[8]; // Get the user status from the eighth column of the event data
        const location = event[9]; // Get the location from the ninth column of the event data
        const sendInvites = event[10]; // Get the send invites option from the tenth column of the event data
        let deleteOpt = event[11]; // Get the delete option from the eighth column of the event data

        let myStatus; // Initialize the myStatus variable to store the user's status
        switch (myStatusStr) { // Determine the user's status based on the string value
            case "OWNER": // If the status is OWNER, set myStatus to CalendarApp.GuestStatus.OWNER
                myStatus = CalendarApp.GuestStatus.OWNER; // Set the status to OWNER
                break;
            case "INVITED": // If the status is INVITED, set myStatus to CalendarApp.GuestStatus.INVITED
                myStatus = CalendarApp.GuestStatus.INVITED; // Set the status to INVITED
                break;
            case "YES": // If the status is YES, set myStatus to CalendarApp.GuestStatus.YES
                myStatus = CalendarApp.GuestStatus.YES; // Set the status to YES
                break;
            case "NO": // If the status is NO, set myStatus to CalendarApp.GuestStatus.NO
                myStatus = CalendarApp.GuestStatus.NO; // Set the status to NO
                break;
            case "MAYBE": // If the status is MAYBE, set myStatus to CalendarApp.GuestStatus.MAYBE
                myStatus = CalendarApp.GuestStatus.MAYBE; // Set the status to MAYBE
                break;
        }

        if (eventID == undefined && eventTitle == undefined && description == undefined && color == undefined && !(startTime instanceof Date) && !(endTime instanceof Date) && guestsStr == undefined && location == undefined && sendInvites == undefined) return; // If the event ID, title, description, color, start time, end time, guests, location, and send invites are all undefined, skip to the next iteration

        if (deleteOpt) {
            // If the delete option is set, delete the event from the calendar
            DeleteEvent(eventID); // Call the DeleteEvent function to delete the event
            return; // Skip to the next iteration if the event is marked for deletion
        }

        if (isAllDayEv) { // If the event is an all-day event
            if (eventID != '') { // If the event ID is not empty, update the existing event
                UpdateEvent(isAllDayEv, eventID, eventTitle, startTime, endTime, description, color, guests, myStatus, location, sendInvites); // Call the UpdateEvent function to update the existing event
                return; // Skip to the next iteration after updating the event
            }
            CreateEvent(isAllDayEv, eventTitle, startTime, endTime, description, color, guests, location, sendInvites); // If the event ID is empty, create a new all-day event
            return; // Skip to the next iteration after creating the event
        }

        if (eventID != '') { // If the event ID is not empty, update the existing event
            UpdateEvent(isAllDayEv, eventID, eventTitle, startTime, endTime, description, color, guests, myStatus, location, sendInvites); // Call the UpdateEvent function to update the existing event
            return; // Skip to the next iteration after updating the event
        }

        CreateEvent(isAllDayEv, eventTitle, startTime, endTime, description, color, guests, location, sendInvites); // If the event ID is empty, create a new event
    });

    var nextDay = new Date(endDate); // Get the end date from the last event in the data
    nextDay.setDate(nextDay.getDate() + 1); // Increment the end date by one day to ensure the next day's events are included
    AddEventsToSheet({ start: startDate, end: nextDay }); // Call the function to import calendar events to the sheet with the specified start and end dates
}

/**
 * This function retrieves the start and end dates from the events data.
 */
function GetStartEndDates(events) {
    const startArray = events.map(ev => ev[2]); // Extract the start timestamps from the events data
    const minStartTimestamp = Math.min(...startArray); // Get the minimum start timestamp from the startArray
    const maxEndTimestamp = Math.max(...startArray); // Get the maximum end timestamp from the startArray
    const startDate = new Date(minStartTimestamp); // Get the start date from the first event in the data
    const endDate = new Date(maxEndTimestamp); // Get the end date from the last event in the data
    return [startDate, endDate]; // Return an array containing the start and end dates
}

/**
 * Deletes an event from the calendar by its ID.
 */
function DeleteEvent(id) {
    const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer
    const calendar = CalendarApp.getCalendarById(settings.calendarId); // Get the calendar by its ID
    calendar.getEventById(id).deleteEvent(); // Delete the event from the calendar
}

/**
 * Updates an existing event in the calendar or creates a new one if it doesn't exist.
 */
function UpdateEvent(isAllDay, eventID, eventTitle, startTime, endTime, description, color, guests, myStatus, location, sendInvites) {
    try { // Try to update the event in the calendar
        const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer
        const calendar = CalendarApp.getCalendarById(settings.calendarId); // Get the calendar by its ID
        var event = calendar.getEventById(eventID); // Get the event by its ID from the calendar
        if (event) { // If the event exists in the calendar
            if (event.getTitle() !== eventTitle) event.setTitle(eventTitle); // Update the event title if it has changed

            if (isAllDay) { // If the event is an all-day event
                if (startTime < endTime && startTime.getDate() < endTime.getDate()) { // If the start time is before the end time and they are on different days
                    event.setAllDayDates(startTime, endTime); // Set the all-day dates for the event
                } 
                else { // If the start time is after the end time or they are on the same day
                    event.setAllDayDate(startTime); // Set the all-day date for the event
                }
            } else { // If the event is not an all-day event
                if (startTime < endTime) { // If the start time is before the end time
                    event.setTime(startTime, endTime); // Set the start and end time for the event
                }
                else { // If the start time is after the end time
                    var nextDay = new Date(startTime); // Create a new date object for the next day
                    nextDay.setDate(nextDay.getDate() + 1); // Increment the next day by one day
                    event.setTime(startTime, nextDay); // Set the start time and the next day as the end time for the event
                }
            }

            if (event.getDescription() !== description) event.setDescription(description); // Update the event description if it has changed
            if (event.getColor() !== Math.trunc(color)) event.setColor(Math.trunc(color)); // Update the event color if it has changed

            if (guests.length > 0 && guests[0] !== '') { // If there are guests to add to the event
                var calGuestList = event.getGuestList(); // Get the list of guests already invited to the event
                guests.forEach(guest => { // Iterate through each guest in the guests array
                    if (!calGuestList.some(value => value.getEmail() == guest.replace(/\([^)]*\)/, "").trim())) { // If the guest is not already in the calendar guest list
                        event.addGuest(guest.replace(/\([^)]*\)/, "").trim()); // Add the guest to the event
                    }
                });
                calGuestList.forEach(guest => { // Iterate through each guest in the calendar guest list
                    if (!guests.some(value => value.replace(/\([^)]*\)/, "").trim() == guest.getEmail())) { // If the guest is not in the guests array
                        event.removeGuest(guest.getEmail()); // Remove the guest from the event
                    }
                });
            }

            if (event.getMyStatus() !== myStatus) event.setMyStatus(myStatus); // Update the user's status for the event if it has changed
            if (event.getLocation() !== location) event.setLocation(location); // Update the event location if it has changed
            if (sendInvites) event.sendInvites = sendInvites; // Set the send invites option for the event
            isAllDay = isAllDay; // Ensure the isAllDay variable is set correctly
            return; // Exit the function after updating the event
        }
        CreateEvent(isAllDay, eventTitle, startTime, endTime, description, color, guests, location, sendInvites); // If the event does not exist, create a new event with the provided details
    }
    catch (e) { // Catch any errors that occur during the process
        Logger.log('CATCH: ' + e); // Log the error message
        return false; // Return false to indicate that the event could not be updated or created
    }
}

/**
 * Creates a new event in the calendar.
 */
function CreateEvent(isAllDay, eventTitle, startTime, endTime, description, color, guests, location, sendInvites) {
    const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer
    const calendar = CalendarApp.getCalendarById(settings.calendarId); // Get the calendar by its ID
    var nextDay = new Date(startTime); // Create a new date object for the next day
    nextDay.setDate(nextDay.getDate() + 1); // Increment the next day by one day

    var endDate = startTime < endTime ? endTime : nextDay; // Set the end date to the end time if it is after the start time, otherwise set it to the next day

    let options = { // Options for the event
        description: description, // Set the event description
        location: location, // Set the event location
        guests: guests.join(','), // Set the event guests as a comma-separated string
        sendInvites: sendInvites // Set whether to send invites for the event
    };

    let event; // Initialize the event variable
    if (isAllDay) { // If the event is an all-day event
        event = calendar.createAllDayEvent(eventTitle, startTime, endDate, options); // Create an all-day event in the calendar
    }
    else { // If the event is not an all-day event
        event = calendar.createEvent(eventTitle, startTime, endDate, options); // Create a regular event in the calendar
    }
    event.setColor(color); // Set the color of the event
}

/****************************************
            * SETTINGS *
*****************************************/
/**
 * This function displays a prompt dialog to the user to enter their Calendar ID. If the user enters a valid ID, it sets the active sheet's name to that ID and saves it in user properties. If the user cancels or enters an empty ID, it resets the sheet name to 'Calendar Sync Template' and deletes the stored Calendar ID.
 * @see https://developers.google.com/apps-script/reference/spreadsheet/ui#prompt(prompt,initialvalue,buttonset)
 * @see https://developers.google.com/apps-script/reference/properties/properties-service#setproperty(key,value)
 * @see https://developers.google.com/apps-script/reference/properties/properties-service#deleteproperty(key)
 */
function ShowAddCalendarId() {
    sheetsUI = SpreadsheetApp.getUi();  // Get the user interface of the spreadsheet

    var result = sheetsUI.prompt( // Display a prompt dialog to the user
        "Your Calendar ID", // Title of the prompt dialog
        "Please enter the id for the calendar you want to manage.", // Message in the prompt dialog
        sheetsUI.ButtonSet.OK_CANCEL, // Set the button set to OK and Cancel
    );

    // Process the user's response.
    var button = result.getSelectedButton(); // Get the button that the user clicked
    var calId = result.getResponseText(); // Get the text entered by the user in the prompt dialog
    if (button == sheetsUI.Button.OK) { // If the user clicked OK
        if (calId != "") { // If the user entered a Calendar ID
            SpreadsheetApp.getActiveSheet().setName(calId); // Set the active sheet's name to the Calendar ID
            PropertiesService.getUserProperties().setProperty('CALID', calId); // Save the Calendar ID in user properties
            const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer
            SpreadsheetApp.getActive().setSpreadsheetTimeZone(settings.defaultTimeZone); // Set the spreadsheet's time zone to the default time zone of the calendar
        }
        else { // If the user entered an empty Calendar ID
            SpreadsheetApp.getActiveSheet().setName('Calendar Sync Template'); // Reset the active sheet's name to 'Calendar Sync Template'
            PropertiesService.getUserProperties().deleteProperty('CALID'); // Delete the stored Calendar ID from user properties
        }
    }
}

/**
 * Clears the content of the active sheet starting from row 2, column 1 to the last row and last column. This function is typically used to reset the sheet before importing new events or data.
 */
function ClearSheet() {
    const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer
    const sheet = settings.sheet; // Get the active sheet in the spreadsheet
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent(); // Clear the content of the sheet starting from row 2, column 1 to the last row and last column
}

/**
 * Displays a modal dialog with documentation link.
 * @see https://developers.google.com/apps-script/guides/html/communication
 */
function ShowDocumentation() {
    var htmlOutput = HtmlService.createHtmlOutput("<p>For troubleshooting or help please visit the <a href='https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer' target='_blank'>Github Page</a>.</p><p>To report any issues or request a feature please visit <a href='https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer/issues' target='_blank'>this link</a>.</p>"); // Create an HTML output with a link to the documentation
    sheetsUI = SpreadsheetApp.getUi(); // Get the user interface of the spreadsheet
    sheetsUI.showModalDialog(htmlOutput, 'Documentation'); // Show the dialog with the documentation link
}

/**
 * Displays a modal dialog with information about the project and its contributors.
 */
function ShowAbout() {
    var htmlOutput = HtmlService.createHtmlOutput("<p>Google Sheets & Google Calendar Synchronizer helps us to enhance our productivity connecting these two amazing Google tools.</p><p>This project is the vision of Sarah Siqueira and more work can be viewed <a href='https://github.com/sarahcssiqueira'>here</a></p><p>Contributors include: <a href='https://github.com/Mtruvius'>Mtruvius</a></p>"); // Create an HTML output with a link to the documentation
    sheetsUI = SpreadsheetApp.getUi(); // Get the user interface of the spreadsheet
    sheetsUI.showModalDialog(htmlOutput, 'About'); // Show the dialog with the About Info
}

/****************************************
          * HELPER FUNCTIONS *
*****************************************/
/**
 * Checks if a Calendar ID exists in the user properties. If it does not exist or is empty, it displays an error message prompting the user to add their Calendar ID in the settings.
 */
function CalendarIdExists() {
    const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer
    if (settings.calendarId == null || settings.calendarId == "") { // Check if the calendar ID is null or empty
        DisplayError('Please make sure you have added you Calendar ID in settings!'); // Display an error message if the calendar ID is not set
        return false; // Return false to indicate that the calendar ID does not exist
    }
    return true; // Return true to indicate that the calendar ID exists
}

/**
 * Displays an error message in a modal dialog.
 */
function DisplayError(e) {
    const settings = GetSettings(); // Get the settings for the Google Sheets Calendar Synchronizer
    var htmlOutput = HtmlService.createHtmlOutput(`<p style="font-family: 'Poppins';">` + e + '</p>').setWidth(settings.dialogWidth).setHeight(settings.dialogHeight); // Create an HTML output with the error message and set its width and height
    const ui = SpreadsheetApp.getUi(); // Get the user interface of the spreadsheet
    ui.showModalDialog(htmlOutput, 'An error has occured!'); // Show the dialog with the error message
    return; // Exit the function after displaying the error message
}

/**
 * Includes an HTML file and returns its content as a string.
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent(); // Create an HTML output from the specified file and return its content as a string
}













