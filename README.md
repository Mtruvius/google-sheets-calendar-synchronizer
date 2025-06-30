# Google Sheets & Google Calendar Synchronizer

[![Project Status: WIP – Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](https://www.repostatus.org/badges/latest/wip.svg)](https://www.repostatus.org/#wip)
[![Release Version](https://img.shields.io/github/release/sarahcssiqueira/google-sheets-calendar-synchronizer.svg)](https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Support Level](https://img.shields.io/badge/support-may_take_time-yellow.svg)](#support-level)

[Google Sheets](https://www.google.com/sheets/about/) & [Google Calendar](https://workspace.google.com/products/calendar/) Synchronizer helps us to enhance our productivity connecting these two amazing Google tools.

## Requirements

- Google Account
- Some JavaScript Knowledge
- Google Spreadsheets and Google Calendar familiarity

## Usage

- Log in to your Google Account;
- Create a Spreadsheet;
- Go to the tab Extensions > Apps Script menu;

![Google Scripts editor](screenshots/access-google-script-editor.png)

- In the Google App Scripts editor copy and paste the code in the [main.js](https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer/blob/master/main.js) file of this repository;
- Make sure to **add the calendar id you want to manage under settings.
![Google Scripts editor](screenshots/id_add_1.png)![Google Scripts editor](screenshots/id_add_2.png)

-Most issues can be resolved by making sure that your calendar, sheet and apps script are all on the same time zone as per the  images below...
![Google Scripts editor](screenshots/appsScripts_timezone.png)![Google Scripts editor](screenshots/calendar_timezone.png)![Google Scripts editor](screenshots/sheet_timezone.png)
### Sample Spreadsheet

That's a sample of a [spreadsheet already correctly formatted](https://docs.google.com/spreadsheets/d/1hsIxXIkFrDHC8NgcDzTdn_vYz3YZD4BjpjiNrdJPFS0/edit?usp=sharing) to work with the following script.

## References

- [Syncing a Spreadsheet with Google Calendar using Google Scripts to be (or at least try) more productive ](https://dev.to/sarahcssiqueira/syncing-a-spreadsheet-with-google-calendar-using-google-scripts-to-be-or-at-least-try-more-productive-18cc)
- [G Suite Pro Tips](https://workspace.google.com/blog/productivity-collaboration/g-suite-pro-tip-how-to-automatically-add-a-schedule-from-google-sheets-into-calendar)
- [Class Calendar](<https://developers.google.com/apps-script/reference/calendar/calendar?hl=pt-br#createAllDayEvent(String,Date,Object)>)
- [Google Sheets - Use Apps Script to Create Google Calendar Events Automatically](https://www.youtube.com/watch?v=FxxPq2wXcK4)
- [Custom Colors](https://developers.google.com/apps-script/reference/calendar/event-color?hl=pt-br)
- [Custom Menus](https://developers.google.com/apps-script/guides/menus?hl=pt-br)

## License

This project is licensed under the [MIT](https://github.com/sarahcssiqueira/google-sheets-calendar-synchronizer/blob/master/LICENSE) license.
