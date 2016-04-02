 /*
     * Date Format 1.2.3
     * (c) 2007-2009 Steven Levithan <stevenlevithan.com>
     * MIT license
     *
     * Includes enhancements by Scott Trenda <scott.trenda.net>
     * and Kris Kowal <cixar.com/~kris.kowal/>
     *
     * Accepts a date, a mask, or a date and a mask.
     * Returns a formatted version of the given date.
     * The date defaults to the current date/time.
     * The mask defaults to dateFormat.masks.default.
     */

    var dateFormat = function () {
        var    token = /d{1,4}|m{1,4}|yy(?:yy)?|([HhMsTt])\1?|[LloSZ]|"[^"]*"|'[^']*'/g,
            timezone = /\b(?:[PMCEA][SDP]T|(?:Pacific|Mountain|Central|Eastern|Atlantic) (?:Standard|Daylight|Prevailing) Time|(?:GMT|UTC)(?:[-+]\d{4})?)\b/g,
            timezoneClip = /[^-+\dA-Z]/g,
            pad = function (val, len) {
                val = String(val);
                len = len || 2;
                while (val.length < len) val = "0" + val;
                return val;
            };
    
        // Regexes and supporting functions are cached through closure
        return function (date, mask, utc) {
            var dF = dateFormat;
    
            // You can't provide utc if you skip other args (use the "UTC:" mask prefix)
            if (arguments.length == 1 && Object.prototype.toString.call(date) == "[object String]" && !/\d/.test(date)) {
                mask = date;
                date = undefined;
            }
    
            // Passing date through Date applies Date.parse, if necessary
            date = date ? new Date(date) : new Date;
            if (isNaN(date)) throw SyntaxError("invalid date");
    
            mask = String(dF.masks[mask] || mask || dF.masks["default"]);
    
            // Allow setting the utc argument via the mask
            if (mask.slice(0, 4) == "UTC:") {
                mask = mask.slice(4);
                utc = true;
            }
    
            var    _ = utc ? "getUTC" : "get",
                d = date[_ + "Date"](),
                D = date[_ + "Day"](),
                m = date[_ + "Month"](),
                y = date[_ + "FullYear"](),
                H = date[_ + "Hours"](),
                M = date[_ + "Minutes"](),
                s = date[_ + "Seconds"](),
                L = date[_ + "Milliseconds"](),
                o = utc ? 0 : date.getTimezoneOffset(),
                flags = {
                    d:    d,
                    dd:   pad(d),
                    ddd:  dF.i18n.dayNames[D],
                    dddd: dF.i18n.dayNames[D + 7],
                    m:    m + 1,
                    mm:   pad(m + 1),
                    mmm:  dF.i18n.monthNames[m],
                    mmmm: dF.i18n.monthNames[m + 12],
                    yy:   String(y).slice(2),
                    yyyy: y,
                    h:    H % 12 || 12,
                    hh:   pad(H % 12 || 12),
                    H:    H,
                    HH:   pad(H),
                    M:    M,
                    MM:   pad(M),
                    s:    s,
                    ss:   pad(s),
                    l:    pad(L, 3),
                    L:    pad(L > 99 ? Math.round(L / 10) : L),
                    t:    H < 12 ? "a"  : "p",
                    tt:   H < 12 ? "am" : "pm",
                    T:    H < 12 ? "A"  : "P",
                    TT:   H < 12 ? "AM" : "PM",
                    Z:    utc ? "UTC" : (String(date).match(timezone) || [""]).pop().replace(timezoneClip, ""),
                    o:    (o > 0 ? "-" : "+") + pad(Math.floor(Math.abs(o) / 60) * 100 + Math.abs(o) % 60, 4),
                    S:    ["th", "st", "nd", "rd"][d % 10 > 3 ? 0 : (d % 100 - d % 10 != 10) * d % 10]
                };
    
            return mask.replace(token, function ($0) {
                return $0 in flags ? flags[$0] : $0.slice(1, $0.length - 1);
            });
        };
    }();
    
    // Some common format strings
    dateFormat.masks = {
        "default":      "ddd mmm dd yyyy HH:MM:ss",
        shortDate:      "m/d/yy",
        mediumDate:     "mmm d, yyyy",
        longDate:       "mmmm d, yyyy",
        fullDate:       "dddd, mmmm d, yyyy",
        shortTime:      "h:MM TT",
        mediumTime:     "h:MM:ss TT",
        longTime:       "h:MM:ss TT Z",
        isoDate:        "yyyy-mm-dd",
        isoTime:        "HH:MM:ss",
        isoDateTime:    "yyyy-mm-dd'T'HH:MM:ss",
        isoUtcDateTime: "UTC:yyyy-mm-dd'T'HH:MM:ss'Z'"
    };
    
    // Internationalization strings
    dateFormat.i18n = {
        dayNames: [
            "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat",
            "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
        ],
        monthNames: [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
        ]
    };
    
    // For convenience...
    Date.prototype.format = function (mask, utc) {
        return dateFormat(this, mask, utc);
    };

// Use this code for Google Docs, Forms, or new Sheets.
// Display a popup window with instructions.
// Use a backslash to span multiple lines with a string e.g. in the msgBox text.
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Date Range')
      .addItem('Select Dates', 'openDateSelector')
      .addToUi();
  Browser.msgBox('Instructions - Please Read, Mrs. Hess :-)', '1) Click on Date Range in the menu bar.\\n2) Choose Select Dates.\
  \\n3) Pick a start date/time and end date/time.\\n4) Click Submit.\\n5) You should see an alert. Click OK to close it.\
  \\n6) Click Close to exit the Enter Date Range dialog.', Browser.Buttons.OK);
}

function openDateSelector() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Enter Date Range');
}

function processForm(formObject) {
  var startDate = formObject.myStart.replace(/-/g,"/").replace(/T/g," ");
  var endDate = formObject.myEnd.replace(/-/g,"/").replace(/T/g," ");
  var formatStartDate = dateFormat(startDate, "mmmm d, yyyy, h:MM:ss TT"); //use dS for the ordinal of the day instead of the cardinal number
  var formatEndDate = dateFormat(endDate, "mmmm d, yyyy, h:MM:ss TT");
  Logger.log(startDate);
  Logger.log(endDate);
  Logger.log(formatStartDate);
  Logger.log(formatEndDate);
  export_gcal_to_gsheet();

function export_gcal_to_gsheet(){

  function export_income_cal_to_gsheet(){

    var mycal = "[calendaraddress]@group.calendar.google.com";
    var cal = CalendarApp.getCalendarById(mycal);
    var events = cal.getEvents(new Date(startDate), new Date(endDate));

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.clearContents();

    var header = [["Payment Description", "Payment Method", "Amount Due", "Due Date", "Last Updated"]]
    var range = sheet.getRange(5,1,1,5);
    range.setValues(header);
    range.setFontWeight("bold")
    range.setHorizontalAlignment("center");
    range.setBorder(true, true, true, true, null, null);

    for (var i=0;i<events.length;i++) {
    var row=i+6;
    var myformula_placeholder = '';

    var details=[[events[i].getTitle(), events[i].getDescription(), myformula_placeholder, events[i].getStartTime(), events[i].getLastUpdated()]];
    var range=sheet.getRange(row,1,1,5);
    range.setValues(details);
    var cell=sheet.getRange(row,3);
    cell.setFormula('=SPLIT( LOWER(A' +row+ ') ; "abcdefghijklmnopqrstuvwxyz &:" )');
    cell.setNumberFormat("$0.00");
    }

    var dates = sheet.getRange("a1");
    dates.setValue("Date Range:").setHorizontalAlignment("left");
    var bills = sheet.getRange("a2");
    bills.setValue("Total Bills:");
    var due = sheet.getRange("a3");
    due.setValue("Total Due:");
    var bold = sheet.getRange(1, 1, 3, 1);
    bold.setFontWeight("bold");

    sheet.getRange("b1").setValue(formatStartDate).setHorizontalAlignment("left");
    var count=sheet.getRange("b2");
    count.setFormula("=COUNTA(a6:a)");
    count.setHorizontalAlignment("left");
    count.setNumberFormat("0");
    var sum=sheet.getRange("b3");
    sum.setFormula("=SUM(c2:c)");
    sum.setNumberFormat("$0.00");
    sum.setHorizontalAlignment("left");
    sum.setFontColor("#c53929").setFontWeight("bold");
    sheet.getRange("c1").setValue("<<<     to     >>>").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("d1").setValue(formatEndDate).setHorizontalAlignment("left");


  }

//
// Export Google Calendar Events to a Google Spreadsheet
//
// This code retrieves events between 2 dates for the specified calendar.
// It logs the results in the first sheet of the spreadsheet listing the events,
// dates/times, etc.
//
// Reference Websites:
// https://developers.google.com/apps-script/reference/calendar/calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-event

var mycal = "[calendaraddress]@group.calendar.google.com";
var cal = CalendarApp.getCalendarById(mycal);

// Enter beginning and ending date range. -- deprecated 2/16/16 due to implementing jquery datepicker
// var begin_date = ("02-16-2016");
// var end_date = ("February 25, 2016 23:59:59 CST");
// Logger.log(formStart);

// Optional variations on getEvents
// var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"));
// var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"), {search: 'word1'});
// var events = cal.getEvents(new Date("January 1, 2016 00:00:00 CST"), new Date("January 14, 2016 23:59:59 CST"));
var events = cal.getEvents(new Date(startDate), new Date(endDate));

// Get the active spreadsheet and make the first sheet active. Select the first active sheet as a variable. Fixes accidentally overwriting contents of the wrong sheet when running the script with
// another sheet (besides the first sheet) active.
var ss = SpreadsheetApp.getActiveSpreadsheet();
SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
var sheet = SpreadsheetApp.getActiveSheet();

// Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost
// sheet.clearContents();

// Create a header record on the current spreadsheet in cells A1:E1 - Match the number of entries in the "header=" to the last parameter
// of the getRange entry below
// getRange syntax = getRange(row, column, numRows, numColumns)
var header = [["Payment Description", "Payment Method", "Amount Due", "Due Date", "Last Updated"]]
var range = sheet.getRange(5,1,1,5);
range.setValues(header);
range.setFontWeight("bold")
range.setHorizontalAlignment("center");
range.setBorder(true, true, true, true, null, null);
  
// Loop through all calendar events found and write them out starting on calculated ROW 6 (i+6)
for (var i=0;i<events.length;i++) {
var row=i+6;
var myformula_placeholder = '';
// Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
// NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
var details=[[events[i].getTitle(), events[i].getDescription(), myformula_placeholder, events[i].getStartTime(), events[i].getLastUpdated()]];
var range=sheet.getRange(row,1,1,5);
range.setValues(details);

// Writing formulas from scripts requires that you write the formulas separate from non-formulas
// Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
var cell=sheet.getRange(row,3);
cell.setFormula('=SPLIT( LOWER(A' +row+ ') ; "abcdefghijklmnopqrstuvwxyz &:" )');
cell.setNumberFormat("$0.00");

}

// https://developers.google.com/apps-script/reference/spreadsheet/range#setbordertop-left-bottom-right-vertical-horizontal
// setBorder syntax = setBorder(top, left, bottom, right, vertical, horizontal) 
var dates = sheet.getRange("a1");
dates.setValue("Date Range:").setHorizontalAlignment("left");
var bills = sheet.getRange("a2");
bills.setValue("Total Bills:");
var due = sheet.getRange("a3");
due.setValue("Total Due:");
var bold = sheet.getRange(1, 1, 3, 1);
bold.setFontWeight("bold");

//var formattedStartDate = Utilities.formatDate(new Date(startDate), "GMT -0600", "EEEE, MMMM d, yyyy");
//var formattedEndDate = Utilities.formatDate(new Date(endDate), "GMT -0600", "EEEE, MMMM d, yyyy");
//Logger.log(formattedStartDate);
//Logger.log(formattedEndDate);

sheet.getRange("b1").setValue(formatStartDate).setHorizontalAlignment("left");
var count=sheet.getRange("b2");
count.setFormula("=COUNTA(a6:a)");
count.setHorizontalAlignment("left");
count.setNumberFormat("0");
var sum=sheet.getRange("b3");
sum.setFormula("=SUM(c2:c)");
sum.setNumberFormat("$0.00");
sum.setHorizontalAlignment("left");
sum.setFontColor("#c53929").setFontWeight("bold");
sheet.getRange("c1").setValue("<<<     to     >>>").setFontWeight("bold").setHorizontalAlignment("center");
sheet.getRange("d1").setValue(formatEndDate).setHorizontalAlignment("left");

// freeze first five rows
sheet.setFrozenRows(5);

// auto resize columns 1, 2, and 4
sheet.autoResizeColumn(1);
sheet.autoResizeColumn(2);
// sheet.autoResizeColumn(3);
sheet.autoResizeColumn(4);
// sheet.autoResizeColumn(5);

// clean up unused columns
// sheet.deleteColumns(6, 21)

// clean up unused rows: http://stackoverflow.com/questions/33787057/how-to-set-a-sheets-number-of-rows-and-columns-at-creation-time-and-a-word-ab
var sheetRows = sheet.getMaxRows();
var lastRow = sheet.getLastRow();
if (sheetRows > lastRow) {
    sheet.deleteRows(lastRow+1, sheetRows-lastRow);
}

// clean up unused columns
var sheetColumns = sheet.getMaxColumns();
Logger.log(sheetColumns);
var lastColumn = sheet.getLastColumn();
Logger.log(lastColumn);
if (sheetColumns > lastColumn) {
    sheet.deleteColumns(lastColumn+1, sheetColumns-lastColumn);
}
}
}
