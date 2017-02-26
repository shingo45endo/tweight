'use strict';

// Constants
var HEADER_ROW = 1;
var LABEL_ROW  = 2;
var DATA_ROW   = 3;
var HEADER_COL = 1;

var fixedHeaders = ['date', 'datetime', 'category'];
var measurements = {
  1: 'Weight (kg)',
  4: 'Height (meter)',
  5: 'Fat Free Mass (kg)',
  6: 'Fat Ratio (%)',
  8: 'Fat Mass Weight (kg)',
  9: 'Diastolic Blood Pressure (mmHg)',
  10: 'Systolic Blood Pressure (mmHg)',
  11: 'Heart Pulse (bpm)',
  12: 'Temperature',
  54: 'SP02 (%)',
  71: 'Body Temperature',
  73: 'Skin Temperature',
  76: 'Muscle Mass',
  77: 'Hydration',
  88: 'Bone Mass',
  91: 'Pulse Wave Velocity',
};

/**
 * Returns a text label for header.
 * @param {number|string} numstr - A kind of body measurements as a number (1..91) / A property name as a string
 * @return {string} A text label.
 */
function WITHINGSLABEL(numstr) {
  if (!measurements[numstr]) {
    return numstr.charAt(0).toUpperCase() + numstr.slice(1);
  } else {
    return measurements[numstr];
  }
}

/**
 * Gets the sheet that contains the body measurements.
 * If the sheet doesn't exist, creates a new sheet and returns it.
 * @return {Sheet} The sheet that contains the body measurements.
 */
function getSheet() {
  var SHEET_NAME = 'Body Measures';
  var id = SpreadsheetApp.getActiveSpreadsheet().getId();
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    initializeSheet(sheet);
  }

  return sheet;
}

/**
 * Initializes a sheet for storing body measurements.
 * @param {Sheet} sheet - A sheet.
 */
function initializeSheet(sheet) {
  if (!sheet) {
    Logger.log('Invalid argument(s)');
    return;
  }

  // Sets the header row and the label row.
  sheet.getRange(HEADER_ROW, 1, 1, fixedHeaders.length).setValues([fixedHeaders]);
  sheet.getRange(LABEL_ROW, 1, 1, sheet.getLastColumn()).setFormulaR1C1('=WITHINGSLABEL(R[-1]C[0])');

  // Hides the header row and column.
  sheet.hideRows(HEADER_ROW);
  sheet.hideColumns(HEADER_COL);
}

/**
 * Gets the header labels as an array.
 * @param {Sheet} sheet - The sheet that contains the body measurements.
 * @return {Array<string>} The header labels as an array.
 */
function getCurrentHeaders(sheet) {
  if (!sheet) {
    Logger.log('Invalid argument(s)');
    return;
  }

  var rowNum = sheet.getLastRow();
  if (rowNum === 0) {
    initializeSheet(sheet);
  }

  var colNum = sheet.getLastColumn();
  return sheet.getRange(HEADER_ROW, 1, 1, colNum).getValues()[0].map(function(elem) {return String(elem);});
}

/**
 * Updates a sheet with new measurement data.
 * @param {Sheet} sheet - A sheet that contains the body measurements.
 * @param {Array<Object>} items - An array of new measurement data.
 */
function updateSheet(sheet, items) {
  if (!sheet || !items) {
    Logger.log('Invalid argument(s)');
    return;
  }

  // Picks up the all types in the items.
  var typeSet = items.reduce(function(obj, item) {
    for (var key in item) {
      if (/^\d+$/.test(key)) {
        obj[key] = true;
      }
    }
    return obj;
  }, {});

  // Makes an array of types.
  var headers = getCurrentHeaders(sheet);
  var types = Object.keys(typeSet);
  types.sort(function(a, b) {return Number(a) - Number(b);});
  types.forEach(function(type) {
    if (headers.indexOf(type) === -1) {
      headers.push(type);
    }
  });

  // Updates the header row of the sheet.
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  initializeSheet(sheet);

  // Gets the keys of the stored data for checking whether the each item has already been stored or not.
  var rowNum = sheet.getLastRow();
  var primaryKeys = sheet.getRange(1, HEADER_COL, rowNum, 1).getValues().map(function(aoa) {return aoa[0];});
  primaryKeys = primaryKeys.slice(DATA_ROW - 1);

  // Appends the new items to the sheet.
  var cells = [];
  items.forEach(function(item) {
    if (primaryKeys.indexOf(item.date) !== -1) {
      return;
    }
    cells.push(headers.map(function(header) {return (item[header] !== undefined) ? item[header] : '';}));
  });
  if (cells.length > 0) {
    sheet.getRange(rowNum + 1, 1, cells.length, headers.length).setValues(cells);
  }
}

/**
 * Reads the measurement data from a sheet.
 * @param {Sheet} sheet - A sheet that contains the body measurements.
 * @return {Array<Object>} An array of measurement data.
 */
function readDataFromSheet(sheet) {
  var items = [];
  if (!sheet) {
    Logger.log('Invalid argument(s)');
    return items;
  }

  var rowNum = sheet.getLastRow();
  var colNum = sheet.getLastColumn();

  var cells = sheet.getRange(1, 1, rowNum, colNum).getValues();
  var headers = cells.shift();
  cells.shift();  // for the label row

  cells.forEach(function(cell) {
    var item = cell.reduce(function(obj, elem, i) {
      if (headers[i] && elem) {
        obj[headers[i]] = elem;
      }
      return obj;
    }, {});
    items.push(item);
  });

  return items;
}

/**
 * Convert a data from Withings API to an array for storing to sheets.
 * @param {Object} data - A JSON data from Withings API.
 * @return {Array<Object>} An array of measurement data.
 */
function convertWithingsData(data) {
  var items = [];
  if (!data || (data.status && data.status !== 0) || !data.body || !data.body.measuregrps) {
    return items;
  }

  var timezone = data.body.timezone || 'GMT';
  data.body.measuregrps.forEach(function(measuregrp) {
    if (!measuregrp.date || !measuregrp.measures) {
      return;
    }

    var item = measuregrp.measures.reduce(function(obj, measure) {
      if (!measure.type || !measure.value || !measure.unit) {
        return obj;
      }
      if (measure.unit < 0) {
        obj[measure.type] = measure.value / Math.pow(10, -measure.unit);
      } else {
        obj[measure.type] = measure.value * Math.pow(10, measure.unit);
      }
      return obj;
    }, {});

    item.date = measuregrp.date;
    item.datetime = Utilities.formatDate(new Date(measuregrp.date * 1000), timezone, 'yyyy/MM/dd HH:mm:ss');
    item.category = measuregrp.category;

    items.push(item);
  });

  items.sort(function(a, b) {return a.date - b.date;});

  return items;
}

var withings = new WithingsWebService(
  PropertiesService.getScriptProperties().getProperty('withingsConsumerKey'),
  PropertiesService.getScriptProperties().getProperty('withingsConsumerSecret'),
  withingsAuthCallback);

/**
 * Handles OAuth callback for Withings API.
 * @param {Object} request - The request data recieved from the callback function.
 */
function withingsAuthCallback(request) {
  return withings.authCallback(request);
}

/**
 * Triggers to get and store data with Withings API.
 * (This function is assumed to be called by Time-driven triggers.)
 */
function triggerStoringDataFromWithingsApi() {
  // Creates the OAuth1 service for Withings API.
  var service = withings.getService();
  if (!service.hasAccess()) {
    Logger.log('Needs to authorize.');
    withings.logAuthorizationUrl();
    withings.logCallbackUrl();
    return;
  }

  // Accesses to the Withings API to get body measure data.
  var userid = PropertiesService.getScriptProperties().getProperty('userid');
  var url = 'https://wbsapi.withings.net/measure?action=getmeas&userid=' + userid;
  var response = service.fetch(url);
  if (response.getResponseCode() >= 400) {
    Logger.log('Something wrong with HTTP response:');
    Logger.log(response);
    return;
  }

  // Stores the obtained body measure data into the sheet.
  var result = JSON.parse(response.getContentText());
  if (result.status !== 0) {
    Logger.log('Something wrong with Withings API: %s', response.getContentText());
    return;
  }
  var items = convertWithingsData(result);
  updateSheet(getSheet(), items);
}

/**
 * Makes a chart.
 * @param {Array<Object>} items - An array of measurement data.
 * @param {number} day - Number of days as a period of chart.
 * @return {Chart} A generated line chart.
 */
function makeChart(items, day) {
  if (!items) {
    return null;
  }
  day = day || 7;

  // Decides the time span of the chart.
  var latest = Math.max.apply(null, items.map(function(item) {return item.date || 0;}));
  var end = Math.ceil((latest + 1) / (60 * 60)) * 60 * 60;
  var begin = end - day * 24 * 60 * 60;

  // Initializes the cells.
  var rows = new Array(24 * day);
  for (var i = 0; i < rows.length; i++) {
    rows[i] = [new Date((begin + i * 60 * 60) * 1000), null];
  }

  // Inserts the data into the cells.
  items.forEach(function(item) {
    if (item.date < begin) {
      return;
    }
    var index = Math.floor((item.date - begin) / (60 * 60));
    rows[index] = [new Date(item.date * 1000), item[1]];
  });

  // Adjusts the timezone of the Date objects.
  // Charts class seems to display Date values in PST. So, needs to adjust the values only for display.
  // '-8 * 60 * 60' means the time difference between PST and UTC.
  // 'getTimezoneOffset() * 60' means the time difference between UTC and the local time.
  // TODO: Support the summer time.
  var offset = -8 * 60 * 60 + (new Date()).getTimezoneOffset() * 60;
  rows.forEach(function(row) {row[0].setTime(row[0].getTime() - offset * 1000);});

  // Makes a DataTable from the cells.
  var data = Charts.newDataTable()
  .addColumn(Charts.ColumnType.DATE, 'Date')
  .addColumn(Charts.ColumnType.NUMBER, measurements[1]);
  rows.forEach(function(row) {data.addRow(row);});
  data.build();

  // Makes parameters for LineChart.
  var maxValue = Math.ceil(Math.max.apply(null, rows.map(function(row) {return row[1] || 0;})) + 0.1);
  var minValue = Math.floor(Math.min.apply(null, rows.map(function(row) {return row[1] || 1000;})) - 0.1);
  var ticks = [];
  for (var v = minValue; v <= maxValue; v += 0.5) { // 'step' should be a precise value in binary representation. (ex. 1, 0.5, 0.25, ...)
    ticks.push(v);
  }

  // Makes a chart from the DataTable.
  var chart = Charts.newLineChart()
  .setDataTable(data)
  .setDimensions(512, 256)
  .setCurveStyle(Charts.CurveStyle.SMOOTH)
  .setPointStyle(Charts.PointStyle.MEDIUM)
  .setLegendPosition(Charts.Position.NONE)
  .setOption('interpolateNulls', true)
  .setOption('hAxis', {format: 'M/d'})
  .setOption('vAxis', {
    maxValue: maxValue,
    minValue: minValue,
    ticks: ticks,
    textPosition: 'none',
  })
  .setOption('chartArea', {width: '90%', height: '82.5%'})
  .build();

  return chart;
}

/**
 * Makes a text for tweet.
 * @param {Array<Object>} items - An array of measurement data.
 * @return {string} A text for tweet.
 */
function makeText(items) {
  if (!items) {
    return null;
  }

  // Picks up 2 items for diff.
  var dates = items.map(function(item) {return item.date || 0;});
  var latest = Math.max.apply(null, dates);
  var prev = dates.reduce(function(nearest, current) {
    if (current === latest) {
      return nearest;
    }
    var diffN = Math.abs(24 * 60 * 60 - Math.abs(latest - nearest));
    var diffC = Math.abs(24 * 60 * 60 - Math.abs(latest - current));
    return (diffC < diffN) ? current : nearest;
  }, 0);
  var latestIndex = dates.indexOf(latest);
  var prevIndex = dates.indexOf(prev);

  // Begins making a text.
  var text = '';

  // Adds changes in weight.
  var diffWeight = items[latestIndex]['1'] - items[prevIndex]['1']; // '1' means 'Weight (kg)'. 
  if (diffWeight > 0) {
    text += Utilities.formatString('\ud83d\udcc8 %+.1fkg ⤴', diffWeight);
  } else if (diffWeight < 0) {
    text += Utilities.formatString('\ud83d\udcc9 %+.1fkg ⤵', diffWeight);
  } else {
    text += '\ud83d\udcca ±0.0kg ➡';
  }
  text += '\n';

  // Adds the period of time.
  var clocks = ['\ud83d\udd5b', '\ud83d\udd67', '\ud83d\udd50', '\ud83d\udd5c', '\ud83d\udd51', '\ud83d\udd5d', '\ud83d\udd52', '\ud83d\udd5e', '\ud83d\udd53', '\ud83d\udd5f', '\ud83d\udd54', '\ud83d\udd60', '\ud83d\udd55', '\ud83d\udd61', '\ud83d\udd56', '\ud83d\udd62', '\ud83d\udd57', '\ud83d\udd63', '\ud83d\udd58', '\ud83d\udd64', '\ud83d\udd59', '\ud83d\udd65', '\ud83d\udd5a', '\ud83d\udd66'];
  var offset = (new Date()).getTimezoneOffset() * 60;
  var objs = [latest, prev].map(function(timestamp) {
    var date = new Date((timestamp - offset) * 1000);
    return {
      date: date,
      datestr: Utilities.formatDate(date, 'GMT', 'MMM. d, h a'),
      clock: clocks[Math.floor((date.getUTCHours() * 60 + date.getUTCMinutes()) / 30) % clocks.length]
    };
  });
  var diffHour = Math.round((objs[0].date.getTime() - objs[1].date.getTime()) / (1000 * 60 * 60));
  text += Utilities.formatString('%s %s ←(%dh)← %s %s', objs[0].clock, objs[0].datestr, diffHour, objs[1].clock, objs[1].datestr);
  text += '\n';

  // Adds a hashtag.
  text += '#tweight';

  // Trims the text within 140 chars. (Very easy way)
  text = text.slice(0, 140);

  return text;
}

var twitter = new TwitterWebService(
  PropertiesService.getScriptProperties().getProperty('twitterConsumerKey'),
  PropertiesService.getScriptProperties().getProperty('twitterConsumerSecret'),
  twitterAuthCallback);

/**
 * Handles OAuth callback for Twitter API.
 * @param {Object} request - The request data recieved from the callback function.
 */
function twitterAuthCallback(request) {
  return twitter.authCallback(request);
}

/**
 * Triggers to tweet a text with a chart.
 * (This function is assumed to be called by Time-driven triggers.)
 */
function triggerTweeting() {
  // Creates the OAuth1 service for Twitter API.
  var service = twitter.getService();
  if (!service.hasAccess()) {
    Logger.log('Needs to authorize.');
    twitter.logAuthorizationUrl();
    twitter.logCallbackUrl();
    return;
  }

  // Gets historical data from the sheet.
  var items = readDataFromSheet(getSheet());

  // If there's not new measurement data within 24 hours, skip this tweet.
  var latest = Math.max.apply(null, items.map(function(item) {return item.date || 0;}));
  if (Date.now() - (new Date(latest * 1000)).getTime() > 24 * 60 * 60 * 1000) {
    Logger.log();
    return;
  }

  // Tweets the latest change of weight.
  var chart = makeChart(items, 7);
  var text = makeText(items);
  twitter.tweetWithMedia(text, chart.getAs('image/png'));
}

/**
 * Triggers to store data and tweet.
 * (This function is assumed to be called by Time-driven triggers.)
 */
function triggerStoringAndTweeting() {
  triggerStoringDataFromWithingsApi();
  triggerTweeting();
}
