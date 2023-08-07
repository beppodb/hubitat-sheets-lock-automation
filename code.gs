// Hubitat-Sheets Lock Automation for Short-Term Rentals
// (c) 2023 Dan Bedard 
// This code is licensed under MIT license (see LICENSE.txt for details)
// Provided as is, without warranty. Use at your own risk. 
// https://github.com/beppodb/hubitat-sheets-lock-automation

// Summary of Instructions
// More detailed instructions in README.md

// 1. Set up your locks in Hubitat.
// 2. Install "Maker API" with remote access on your Hubitat.
// 3. Create a new Google Sheet.
// 4. Rename the individual tab to match the name of your lock device.
// 5. Add additional tabs for additional locks.
// 6. Create an Apps Script by clicking "Exetensions --> Apps Script"
// 7. Copy and paste this code over the Code.gs file.
// 8. Replace HUBITAT_ACCESS_TOKEN and HUBITAT_URL_STUB with entries from 
// 9. In the PROPERTIES constant below, enter the Google Sheets tab names in the 'sheetname' field and
//    links to your .ics calendars in the 'link' field.
// 10. Save this script.
// 11. Go back to the Google Sheet and refresh.
// 12. On each lock tab, clieck "Rental Lock Automator --> Initialize Sheet"
// 13. Add permanent codes to your locks with the word "permanent" (all lower case) in the "Type" column.
// 14. [Test various functions by running them in Apps Script: updateCalendars() and updateLockStatus() and reviewSheetsForChanges()]
// 15. Once you're confident everything works, add automation triggers: Run everyMinute() on a time-based trigger. Run onChange() on edits.

 
const HUBITAT_ACCESS_TOKEN = "[ADD YOUR ACCESS TOKEN HERE]"
const HUBITAT_URL_STUB = `https://cloud.hubitat.com/api/[ADD YOUR URL HERE]`
const HUBITAT_URL = `${HUBITAT_URL_STUB}/all?access_token=${HUBITAT_ACCESS_TOKEN}`

const HEADER_ROW = ['Slot','Name','Code','Begin','End','Type','Reference']
const NUMBER_FORMATS = ["@","@","@","MM/dd/yy hh:mm","MM/dd/yy hh:mm","@","@"]
const NUM_SLOTS = 30;

const CALENDAR_UPDATE_MINUTES = 30

const PROPERTIES = [
  {
    'sheetname' : 'ADD YOUR FIRST PROPERTY NAME HERE',
    'link' : 'ADD YOUR URL HERE',
    'start_time_offset_hours' : 4,
    'end_time_offset_hours' : 4
  },
    {
    'sheetname' : 'ADD YOUR SECOND PROPERTY NAME HERE',
    'link' : 'ADD YOUR URL HERE',
    'start_time_offset_hours' : 4,
    'end_time_offset_hours' : 4
  }
]

const VRBO_EMAIL = "ENTER THE EMAIL ADDRESS YOU RECEIVE VRBO EMAILS FROM HERE"

var scriptPrp = PropertiesService.getScriptProperties();
var lastChangeTimestamp = '0';
var lastCalendarUpdate = '0';
var changed = 'false';

// Add a custom menu to the active document.
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Rental Lock Automator')
      .addItem('Initialize Sheet', 'initializeThisSheet')
      .addToUi();
}

function onChange(e) {
  lastChangeTimestamp = new Date().getTime();
  scriptPrp.setProperty('lastChangeTimestamp', lastChangeTimestamp.toString());
  scriptPrp.setProperty('changed','true');
}

function updateCalendars() {
  for (var property of PROPERTIES) {
    Logger.log("Merging appointments for " + property.sheetname);
    mergeAppointments(property);
  }
  lastCalendarUpdate = new Date().getTime();
  scriptPrp.setProperty('lastCalendarUpdate', lastCalendarUpdate.toString());
  scriptPrp.setProperty('lastChangeTimestamp', lastCalendarUpdate.toString());
  scriptPrp.setProperty('changed','true');
}

function setThingGetThing() {
   var newtime = scriptPrp.getProperty('lastCalendarUpdate')
    Logger.log('Read ' + newtime)
    var time = new Date().getTime();
    Logger.log("Writing " + time)
    scriptPrp.setProperty('lastCalendarUpdate', time.toString());
    var newtime = scriptPrp.getProperty('lastCalendarUpdate')
    Logger.log('Read ' + newtime)
    new Date()
}

function everyMinute() {
  var currentTime = new Date().getTime();
  var oneMinuteAgo = currentTime - 1 * 60 * 1000;
  var lastChangeTimestamp = scriptPrp.getProperty('lastChangeTimestamp')
  var lastCalendarUpdate = scriptPrp.getProperty('lastCalendarUpdate')
  var changed = scriptPrp.getProperty('changed')

  Logger.log("Current Time: " + currentTime + "\nLast Change: " + lastChangeTimestamp + "\nLast Calendar Update: " + lastCalendarUpdate + "\nChanged: " + changed)

  if(!lastCalendarUpdate || (lastCalendarUpdate < (currentTime - CALENDAR_UPDATE_MINUTES * 60 * 1000))) {
    updateCalendars();
  }
  if(lastChangeTimestamp > oneMinuteAgo) {
    Logger.log("Changes are still too fresh. Exiting...")
    return;
  }
  if(changed == 'true') {
    updateLockStatus();
  }
  reviewSheetsForChanges();
}

function initializeThisSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange(1,1,NUM_SLOTS + 1, HEADER_ROW.length);
  var data = [];
  data.push(HEADER_ROW);
  for (var i = 1; i <= NUM_SLOTS; i++) {
    var row = new Array(HEADER_ROW.length);
    row[HEADER_ROW.indexOf('Slot')] = i.toString();
    data.push(row);
  }
  dataRange.setValues(data);
  for (var i = 1; i <= NUMBER_FORMATS.length; i++) {
    var column = sheet.getRange(2,i,NUM_SLOTS,1);
//    Logger.log("Setting column " + i + " to number format " + NUMBER_FORMATS[i-1])
    column.setNumberFormat(NUMBER_FORMATS[i-1]);
  }
}

function mergeAppointments(property) {
  var spreadsheet_name = property.sheetname;
  var calendar_link = property.link;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet_name);
  
  var dataRange = sheet.getRange(2,2,sheet.getLastRow()-1,6)
  var data = dataRange.getValues();
  
  var appointments = fetchAppointments(property); 
  
  var mergedArray = [];
  
  // Process rows from the first array
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var type = row[4];
    var end = row[3];
    var reference = row[5];
    var now = new Date();
    
    if (type === "permanent" || new Date(end) >= now) {
      if (!reference || appointments.some(appointment => appointment[5] === reference)) {
        mergedArray.push(row);
      }
    }
  }

  // Process rows from the second array
  for (var i = 0; i < appointments.length; i++) {
    var appointment = appointments[i];
    var reference = appointment[5];
    
    if (!data.some(row => row[5] === reference)) {
      mergedArray.push(appointment);
    }
  }
  
  // Sort the merged array
  mergedArray.sort(function(a, b) {
    if (a[4] === "permanent" && b[4] !== "permanent") {
      return -1;
    } else if (a[4] !== "permanent" && b[4] === "permanent") {
      return 1;
    } else {
      return new Date(a[2]) - new Date(b[2]);
    }
  });
  if (mergedArray.length > (sheet.getLastRow()-1) ) {
    mergedArray.splice(sheet.getLastRow()-1);
  }
  dataRange.clearContent();
  var newDataRange = sheet.getRange(2,2,mergedArray.length,6);
  newDataRange.setValues(mergedArray);
}

function updateLockStatus() {
  // Perform the HTTPS GET request
  var response = UrlFetchApp.fetch(HUBITAT_URL);
  var jsonString = response.getContentText();
  
  scriptPrp.setProperty('MostRecent',jsonString) 
}

function reviewSheetsForChanges() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var most_recent;
  try {
    most_recent = JSON.parse(scriptPrp.getProperty('MostRecent'));
  } catch(e) {
    updateLockStatus();
    most_recent = JSON.parse(scriptPrp.getProperty('MostRecent'));
  }
  
//  Logger.log("Most recent: " + JSON.stringify(most_recent))

  var locksObject = most_recent;

  var updates_needed = false;
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var lockName = sheet.getName();
    Logger.log("Processing sheet " + lockName);
    updates_needed = compareLockCodes(lockName, locksObject) || updates_needed;
  }
  if (updates_needed) {
    scriptPrp.setProperty('changed','true')
  } else {
    scriptPrp.setProperty('changed','false')
  }
}

function compareLockCodes(lockName, locksObject) {
  var updates_needed = false;
  var device = getDeviceforLock(locksObject,lockName);
  var lockCodesObject = getLockCodesforLock(locksObject, lockName);
//  Logger.log(JSON.stringify(lockCodesObject));

  // Open the Google Sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(lockName);

  // Get the data from the sheet
  if (sheet.getLastRow() < 2) {
     Logger.log("Sheet " + lockName + " has no active rows.")
     return;
  }
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7);
  var data = dataRange.getValues();
  var processedSlots = []; // Array to store the slots processed from lockCodesObject

  // Process each entry in the lockcodes object
  if (lockCodesObject) {

    // Process the lockCodesObject first
    for (var slot in lockCodesObject) {
      processedSlots.push(slot);
      var code = lockCodesObject[slot].code;
      var name = lockCodesObject[slot].name;

      // Find the matching row in the sheet based on the slot number
      var rowData = null;
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] == slot) {
          break;
        }
      }
//      Logger.log("Processing slot " + slot);
      updates_needed = compareRowData(sheet, i, data[i], device, slot, name, code) || updates_needed;
    }
  }

  // Check if there are slots in the sheet that were not processed from lockCodesObject
  for (var i = 0; i < data.length; i++) {
    var slot = data[i][0].toString();
//    Logger.log("Processing updated slot " + slot);
    if (!processedSlots.includes(slot)) {
      updates_needed = compareRowData(sheet, i, data[i], device, slot, "", "")  || updates_needed;
    }
  }
  return updates_needed;
}

function compareRowData (sheet, row, rowData, device, slot, name, code) {
  var updates_needed = false;
  var now = new Date();
  if (rowData) {
    var sheetName = rowData[1].toString();
    var sheetCode = rowData[2].toString();
    var beginDate = new Date(rowData[3]); // Assuming Begin column is at index 3
    var endDate = new Date(rowData[4]); // Assuming End column is at index 4

    // Compare code and name with the sheet entry
    if (sheetCode === "" && code !== sheetCode) {
      updates_needed = true;
      deleteCode(device, slot);
      } else if (beginDate > now) {
        if (code !== "") {
          updates_needed = true;
          deleteCode(device, slot);
        }
      } else if (endDate < now) {
        sheet.getRange(row + 2, 2, 1, 5).clearContent();
        updates_needed = true;
        deleteCode(device, slot);
      } else {
        if (sheetName !== name || sheetCode !== code) {
          // Call the updateCode function if the conditions are met
          updates_needed = true;
          updateCode(device, slot, sheetName, sheetCode);
        }
      }
  } else {
    // Call the deleteCode function if the slot doesn't exist in the sheet
    updates_needed = true;
    deleteCode(device, slot);        
  }
  return updates_needed;
}

function updateCode(device, slot, name, code) {
  Logger.log("Updating " + slot + " on device " + device + " with code " + code + " for name " + name);
  var hubitat_command = `${HUBITAT_URL_STUB}/${device}/setCode/${slot},${code},${name}?access_token=${HUBITAT_ACCESS_TOKEN}`
  var response = UrlFetchApp.fetch(hubitat_command);
  var jsonString = response.getContentText();
  Logger.log("Received response: " + jsonString);
}

function deleteCode(device, slot) {
  Logger.log("Deleting slot " + slot + " on device " + device);
  var hubitat_command = `${HUBITAT_URL_STUB}/${device}/deleteCode/${slot}?access_token=${HUBITAT_ACCESS_TOKEN}`
  var response = UrlFetchApp.fetch(hubitat_command);
  var jsonString = response.getContentText();
  Logger.log("Received response: " + jsonString);
}

function getLockCodesforLock(lockObjects, label) {
  for (var i = 0; i < lockObjects.length; i++) {
//    Logger.log(JSON.stringify(lockObjects[i]));
    if (lockObjects[i].label === label) {
      return JSON.parse(lockObjects[i].attributes.lockCodes);
    }
  }
  return null; // Return null if the label is not found in any lock object
}

function getDeviceforLock(lockObjects, label) {
  for (var i = 0; i < lockObjects.length; i++) {
//    Logger.log(JSON.stringify(lockObjects[i]));
    if (lockObjects[i].label === label) {
      return JSON.parse(lockObjects[i].id);
    }
  }
  return null; // Return null if the label is not found in any lock object
}

function fetchAppointments(property) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(property.sheetname)
  var icsLink = property.link; // Replace with your ICS link

  var response = UrlFetchApp.fetch(icsLink);
  var data = response.getContentText();
  var today = new Date();
  var endDate = new Date(today.getTime() + (90 * 24 * 60 * 60 * 1000));

  var appointments = [];
  
  var lines = data.split('\n');
  var startDate, endDate, uid, description;
  
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].startsWith('BEGIN:VEVENT')) {
      while (i < lines.length && !lines[i].startsWith('END:VEVENT')) {
        if (lines[i].startsWith('DTSTART')) {
          var timeZone = lines[i].match(/TZID=([^:]+)/)[1];
          var dateString = lines[i].match(/\d{8}T\d{6}/)[0]; // Extract date string
          startDate = new Date(Utilities.parseDate(dateString,timeZone,"yyyyMMdd'T'HHmmss").getTime() - property.start_time_offset_hours * 60 * 60 * 1000);          
        }
        if (lines[i].startsWith('DTEND')) {
          var timeZone = lines[i].match(/TZID=([^:]+)/)[1];
          var dateString = lines[i].match(/\d{8}T\d{6}/)[0]; // Extract date string
          endDate = new Date(Utilities.parseDate(dateString,timeZone,"yyyyMMdd'T'HHmmss").getTime() + property.end_time_offset_hours * 60 * 60 * 1000);
        }
        if (lines[i].startsWith('UID')) {
          uid = lines[i].substring(4);
        }
        if (lines[i].startsWith('DESCRIPTION')) {
          description = lines[i].substring(12);
        }
        i++;
      }

      var fields = description.split('\\n');
      var name, phone, code, type;

      if(fields && fields.length > 5) {
        name = fields[5];
      }

      if(uid.startsWith('airbnb')) {
        if(fields && fields.length > 6) {
          phone = fields[6].substring(6)
          code = phone.substring(phone.length - 4)
          type = "airbnb";
        }
      }
      if(uid.startsWith('homeaway')) {
        phone = getPhoneFromEmail(name, startDate, endDate);
        code = phone.substring(phone.length - 4)
        type = "vrbo"
      }
      
      if (startDate && startDate >= today && startDate <= endDate) {
        console.log("Adding appointment: " + startDate + " " + endDate + " " + " " + name + " " + phone + " " + code)
        appointments.push([type + ":" + Utilities.formatDate(startDate,timeZone,"yyyyMMdd"), code, startDate, endDate, type, type + ": " + name]);
      }
      
      startDate = endDate = description = null;
    }
  }
  return appointments;
}

function getPhoneFromEmail(name, startDate, endDate) {
  var query = "from:" + VRBO_EMAIL + " subject:(\"" + name + ": " + Utilities.formatDate(startDate, "GMT", "MMM d") + " - " + Utilities.formatDate(endDate, "GMT", "MMM d, yyyy") + "\")";
  Logger.log("Querying email for: " + query)
  try {
    var threads = GmailApp.search(query, 0, 1); // Search for the latest email with the specified subject

    if (threads.length > 0) {
      var messages = threads[0].getMessages();
      var emailContent = messages[0].getPlainBody();
  //    Logger.log(emailContent)
      var phoneNumber = extractPhoneNumberFromEmail(emailContent);

      return phoneNumber;
    } else {
      return null; // Email not found
    }
  } catch (error) {
    Logger.log("Error: " + error);
    return null;
  }
}

function extractPhoneNumberFromEmail(emailContent) {
  // Search for the pattern "Traveler Phone:" followed by a phone number
  var pattern = /Traveler Phone:[\s]*(.*?)[\r\n]/
  var match = emailContent.match(pattern);

  if (match && match[1]) {
    return match[1].replace(/\D+/g, ''); // Return the extracted phone number without any non-digits
  } else {
    return null; // Phone number not found
  }
}
